import { describe, expect, it } from "vitest";
import {
  buildFileAttachment,
  detectMimeType,
  extractGuidFromETag,
  type FileUploadResult,
  formatFileSize,
} from "../file-upload.js";

describe("detectMimeType", () => {
  it("should detect common document types", () => {
    expect(detectMimeType("report.pdf")).toBe("application/pdf");
    expect(detectMimeType("doc.docx")).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    expect(detectMimeType("sheet.xlsx")).toBe(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    expect(detectMimeType("archive.zip")).toBe("application/zip");
    expect(detectMimeType("data.csv")).toBe("text/csv");
    expect(detectMimeType("config.json")).toBe("application/json");
  });

  it("should detect image types", () => {
    expect(detectMimeType("photo.png")).toBe("image/png");
    expect(detectMimeType("photo.jpg")).toBe("image/jpeg");
    expect(detectMimeType("photo.jpeg")).toBe("image/jpeg");
    expect(detectMimeType("animation.gif")).toBe("image/gif");
  });

  it("should handle case-insensitive extensions", () => {
    expect(detectMimeType("FILE.PDF")).toBe("application/pdf");
    expect(detectMimeType("IMAGE.PNG")).toBe("image/png");
  });

  it("should return octet-stream for unknown types", () => {
    expect(detectMimeType("file.xyz")).toBe("application/octet-stream");
    expect(detectMimeType("noext")).toBe("application/octet-stream");
  });

  it("should handle full paths", () => {
    expect(detectMimeType("/home/user/docs/report.pdf")).toBe("application/pdf");
    expect(detectMimeType("/tmp/data.csv")).toBe("text/csv");
  });
});

describe("extractGuidFromETag", () => {
  it("should extract GUID from standard eTag format", () => {
    const eTag = '"{B5765B1F-4D42-4C53-91A2-D49A14B0C8C3},2"';
    expect(extractGuidFromETag(eTag)).toBe("B5765B1F-4D42-4C53-91A2-D49A14B0C8C3");
  });

  it("should handle eTag with different version numbers", () => {
    const eTag = '"{ABCDEF01-2345-6789-ABCD-EF0123456789},15"';
    expect(extractGuidFromETag(eTag)).toBe("ABCDEF01-2345-6789-ABCD-EF0123456789");
  });

  it("should handle fallback for non-standard eTag", () => {
    const result = extractGuidFromETag("some-random-string");
    expect(result).toBeTruthy();
  });
});

describe("buildFileAttachment", () => {
  it("should build correct reference attachment structure", () => {
    const uploadResult: FileUploadResult = {
      webUrl: "https://sharepoint.com/file.pdf",
      attachmentId: "B5765B1F-4D42-4C53-91A2-D49A14B0C8C3",
      fileName: "report.pdf",
      fileSize: 1024,
      mimeType: "application/pdf",
    };

    const attachments = buildFileAttachment(uploadResult);

    expect(attachments).toEqual([
      {
        id: "B5765B1F-4D42-4C53-91A2-D49A14B0C8C3",
        contentType: "reference",
        contentUrl: "https://sharepoint.com/file.pdf",
        name: "report.pdf",
      },
    ]);
  });

  it("should return an array with exactly one attachment", () => {
    const uploadResult: FileUploadResult = {
      webUrl: "https://example.com/file.docx",
      attachmentId: "test-guid",
      fileName: "document.docx",
      fileSize: 2048,
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    };

    const attachments = buildFileAttachment(uploadResult);
    expect(attachments).toHaveLength(1);
    expect(attachments[0].contentType).toBe("reference");
  });
});

describe("formatFileSize", () => {
  it("should format bytes", () => {
    expect(formatFileSize(500)).toBe("500 B");
    expect(formatFileSize(0)).toBe("0 B");
  });

  it("should format kilobytes", () => {
    expect(formatFileSize(1536)).toBe("1.5 KB");
    expect(formatFileSize(1024)).toBe("1.0 KB");
  });

  it("should format megabytes", () => {
    expect(formatFileSize(5 * 1024 * 1024)).toBe("5.0 MB");
  });

  it("should format gigabytes", () => {
    expect(formatFileSize(2 * 1024 * 1024 * 1024)).toBe("2.0 GB");
  });
});
