import { promises as fs } from "node:fs";
import { basename, extname } from "node:path";
import type { GraphService } from "../services/graph.js";

/** Simple upload threshold (4 MB) */
const SIMPLE_UPLOAD_MAX_SIZE = 4 * 1024 * 1024;

/** Upload session chunk size — must be a multiple of 320 KiB */
const UPLOAD_CHUNK_SIZE = 320 * 1024 * 10; // 3.2 MB

/** Extension → MIME type map for common file types */
const MIME_TYPES: Record<string, string> = {
  ".pdf": "application/pdf",
  ".doc": "application/msword",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xls": "application/vnd.ms-excel",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".ppt": "application/vnd.ms-powerpoint",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".zip": "application/zip",
  ".7z": "application/x-7z-compressed",
  ".rar": "application/vnd.rar",
  ".tar": "application/x-tar",
  ".gz": "application/gzip",
  ".txt": "text/plain",
  ".csv": "text/csv",
  ".json": "application/json",
  ".xml": "application/xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".webp": "image/webp",
  ".bmp": "image/bmp",
  ".mp4": "video/mp4",
  ".mp3": "audio/mpeg",
  ".wav": "audio/wav",
  ".html": "text/html",
  ".htm": "text/html",
  ".css": "text/css",
  ".js": "application/javascript",
  ".ts": "application/typescript",
  ".py": "text/x-python",
  ".md": "text/markdown",
  ".log": "text/plain",
};

export interface FileUploadResult {
  webUrl: string;
  attachmentId: string;
  fileName: string;
  fileSize: number;
  mimeType: string;
}

/** Graph API response from a DriveItem upload (simple PUT or final chunk). */
type DriveItemUploadResponse = {
  webUrl?: string;
  eTag?: string;
};

/** Graph API response when creating a resumable upload session. */
type UploadSessionResponse = {
  uploadUrl?: string;
};

/** Graph API response for the channel filesFolder endpoint. */
type ChannelFilesFolderResponse = {
  id?: string;
  parentReference?: { driveId?: string };
};

/**
 * Detect MIME type from file extension.
 */
export function detectMimeType(filePath: string): string {
  const ext = extname(filePath).toLowerCase();
  return MIME_TYPES[ext] || "application/octet-stream";
}

/**
 * Extract attachment GUID from the eTag returned by Microsoft Graph.
 * eTag format: `"{GUID},version"` → extracts the GUID portion.
 */
export function extractGuidFromETag(eTag: string): string {
  const match = eTag.match(/\{([^}]+)\}/);
  if (match) {
    return match[1];
  }
  const [rawId] = eTag.split(",");
  return rawId.replace(/["{}]/g, "") || eTag;
}

/**
 * Read a local file and return its contents as a Buffer.
 */
export async function readLocalFile(filePath: string): Promise<{ buffer: Buffer; size: number }> {
  const buffer = await fs.readFile(filePath);
  return { buffer, size: buffer.length };
}

/**
 * Simple upload for files ≤ 4 MB.
 * PUT /drives/{driveId}/items/{parentItemId}:/{fileName}:/content
 */
async function simpleUpload(
  graphService: GraphService,
  driveId: string,
  parentItemId: string,
  remotePath: string,
  fileBuffer: Buffer,
  mimeType: string
): Promise<{ webUrl: string; eTag: string }> {
  const client = await graphService.getClient();
  const response = (await client
    .api(`/drives/${driveId}/items/${parentItemId}:/${remotePath}:/content`)
    .header("Content-Type", mimeType)
    .put(fileBuffer)) as DriveItemUploadResponse;
  if (!response?.webUrl || !response?.eTag) {
    throw new Error("Upload failed: response did not contain webUrl/eTag");
  }
  return { webUrl: response.webUrl, eTag: response.eTag };
}

/**
 * Upload session for files > 4 MB.
 * Creates a resumable upload session and sends the file in 3.2 MB chunks.
 */
async function uploadLargeFile(
  graphService: GraphService,
  driveId: string,
  parentItemId: string,
  remotePath: string,
  fileBuffer: Buffer
): Promise<{ webUrl: string; eTag: string }> {
  const client = await graphService.getClient();

  const session = (await client
    .api(`/drives/${driveId}/items/${parentItemId}:/${remotePath}:/createUploadSession`)
    .post({
      item: {
        "@microsoft.graph.conflictBehavior": "rename",
      },
    })) as UploadSessionResponse;

  if (!session?.uploadUrl) {
    throw new Error("Upload failed: upload session did not return uploadUrl");
  }
  const uploadUrl: string = session.uploadUrl;
  const fileSize = fileBuffer.length;

  let offset = 0;
  let lastResponse: Response | null = null;

  while (offset < fileSize) {
    const chunkEnd = Math.min(offset + UPLOAD_CHUNK_SIZE, fileSize);
    const chunk = fileBuffer.subarray(offset, chunkEnd);
    const contentRange = `bytes ${offset}-${chunkEnd - 1}/${fileSize}`;

    lastResponse = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.length),
        "Content-Range": contentRange,
      },
      body: new Uint8Array(chunk),
    });

    if (!lastResponse.ok) {
      const errorText = await lastResponse.text();
      throw new Error(`Upload chunk failed (${lastResponse.status}): ${errorText}`);
    }

    // Drain intermediate 202 response bodies to free resources
    if (lastResponse.status === 202) {
      await lastResponse.text();
    }

    offset = chunkEnd;
  }

  if (!lastResponse) {
    throw new Error("Upload failed: no response received");
  }
  const finalResult = await lastResponse.json();
  if (!finalResult?.webUrl || !finalResult?.eTag) {
    throw new Error("Upload failed: final response did not contain file metadata");
  }
  return { webUrl: finalResult.webUrl, eTag: finalResult.eTag };
}

/**
 * Upload a file to a Teams channel's SharePoint folder.
 */
export async function uploadFileToChannel(
  graphService: GraphService,
  teamId: string,
  channelId: string,
  filePath: string,
  customFileName?: string
): Promise<FileUploadResult> {
  const client = await graphService.getClient();

  // Get the channel's SharePoint drive and folder IDs
  const filesFolder = (await client
    .api(`/teams/${teamId}/channels/${channelId}/filesFolder`)
    .get()) as ChannelFilesFolderResponse;
  if (!filesFolder?.parentReference?.driveId || !filesFolder?.id) {
    throw new Error("Failed to resolve channel drive/folder IDs");
  }
  const driveId: string = filesFolder.parentReference.driveId;
  const channelFolderId: string = filesFolder.id;

  const fileName = customFileName || basename(filePath);
  const mimeType = detectMimeType(filePath);
  const { buffer, size } = await readLocalFile(filePath);

  const encodedName = encodeURIComponent(fileName);
  const uploadResult =
    size <= SIMPLE_UPLOAD_MAX_SIZE
      ? await simpleUpload(graphService, driveId, channelFolderId, encodedName, buffer, mimeType)
      : await uploadLargeFile(graphService, driveId, channelFolderId, encodedName, buffer);

  return {
    webUrl: uploadResult.webUrl,
    attachmentId: extractGuidFromETag(uploadResult.eTag),
    fileName,
    fileSize: size,
    mimeType,
  };
}

/**
 * Upload a file to OneDrive's "Microsoft Teams Chat Files" folder for chat messages.
 */
export async function uploadFileToChat(
  graphService: GraphService,
  filePath: string,
  customFileName?: string
): Promise<FileUploadResult> {
  const client = await graphService.getClient();

  const fileName = customFileName || basename(filePath);
  const mimeType = detectMimeType(filePath);
  const { buffer, size } = await readLocalFile(filePath);

  const driveResponse = (await client.api("/me/drive").get()) as { id?: string };
  if (!driveResponse?.id) {
    throw new Error("Failed to resolve user drive ID");
  }
  const driveId: string = driveResponse.id;

  const remotePath = `${encodeURIComponent("Microsoft Teams Chat Files")}/${encodeURIComponent(fileName)}`;
  const uploadResult =
    size <= SIMPLE_UPLOAD_MAX_SIZE
      ? await simpleUpload(graphService, driveId, "root", remotePath, buffer, mimeType)
      : await uploadLargeFile(graphService, driveId, "root", remotePath, buffer);

  return {
    webUrl: uploadResult.webUrl,
    attachmentId: extractGuidFromETag(uploadResult.eTag),
    fileName,
    fileSize: size,
    mimeType,
  };
}

/**
 * Build the attachments array for a message that references an uploaded file.
 */
export function buildFileAttachment(uploadResult: FileUploadResult): Array<{
  id: string;
  contentType: string;
  contentUrl: string;
  name: string;
}> {
  return [
    {
      id: uploadResult.attachmentId,
      contentType: "reference",
      contentUrl: uploadResult.webUrl,
      name: uploadResult.fileName,
    },
  ];
}

/**
 * Escape special HTML characters in plain text so it can be safely
 * embedded inside an HTML message body.
 */
export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Format a file size in bytes to a human-readable string.
 */
export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}
