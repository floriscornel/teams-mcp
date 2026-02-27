import { describe, expect, it } from "vitest";
import { detectContentType } from "../content-type.js";

describe("detectContentType", () => {
  it("should detect PNG", () => {
    const buffer = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a]);
    expect(detectContentType(buffer)).toBe("image/png");
  });

  it("should detect JPEG", () => {
    const buffer = Buffer.from([0xff, 0xd8, 0xff, 0xe0]);
    expect(detectContentType(buffer)).toBe("image/jpeg");
  });

  it("should detect GIF", () => {
    const buffer = Buffer.from([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]);
    expect(detectContentType(buffer)).toBe("image/gif");
  });

  it("should detect WebP", () => {
    const buffer = Buffer.from([
      0x52, 0x49, 0x46, 0x46, 0x00, 0x00, 0x00, 0x00, 0x57, 0x45, 0x42, 0x50,
    ]);
    expect(detectContentType(buffer)).toBe("image/webp");
  });

  it("should detect BMP", () => {
    const buffer = Buffer.from([0x42, 0x4d, 0x00, 0x00]);
    expect(detectContentType(buffer)).toBe("image/bmp");
  });

  it("should detect PDF", () => {
    const buffer = Buffer.from([0x25, 0x50, 0x44, 0x46]);
    expect(detectContentType(buffer)).toBe("application/pdf");
  });

  it("should return application/octet-stream for unknown content", () => {
    const buffer = Buffer.from([0x00, 0x01, 0x02, 0x03]);
    expect(detectContentType(buffer)).toBe("application/octet-stream");
  });

  it("should return application/octet-stream for very small buffers", () => {
    const buffer = Buffer.from([0x89, 0x50]);
    expect(detectContentType(buffer)).toBe("application/octet-stream");
  });
});
