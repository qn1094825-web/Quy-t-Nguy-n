declare module 'pizzip' {
  export default class PizZip {
    constructor(data?: any, options?: any);
    file(name: string): { asText(): string; asNodeBuffer(): Buffer } | null;
    file(name: string, data: string | Buffer | Uint8Array): this;
    file(regex: RegExp): Array<{ name: string; asText(): string }>;
    generate(options: { type: 'nodebuffer' | 'uint8array' | 'blob' | 'base64'; compression?: string }): any;
  }
}
