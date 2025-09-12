/**
 * Converts a Word document buffer to clean HTML with proper heading detection
 * @param buffer - The Word document buffer
 * @returns Promise<string> - Clean HTML string
 */
export declare function convertDocxToHtml(buffer: Buffer): Promise<string>;
export declare function isValidWordFile(filename: string, mimetype: string): boolean;
//# sourceMappingURL=converter.d.ts.map