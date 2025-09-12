# Word to HTML Converter

A Node.js + Express.js project that converts Microsoft Word documents (.docx, .doc) to clean HTML.

## Project Structure

- `src/` - TypeScript source files
  - `server.ts` - Express server with file upload handling
  - `converter.ts` - Word to HTML conversion logic
- `public/` - Frontend files
  - `index.html` - Web interface with Tailwind CSS
- `dist/` - Compiled JavaScript files
- `package.json` - Project configuration
- `tsconfig.json` - TypeScript configuration

## Quick Start

1. Install dependencies: `npm install`
2. Build project: `npm run build`
3. Start development server: `npm run dev`
4. Open browser: http://localhost:3000

## Features

- Converts Word headings to proper HTML h1-h6 tags
- Preserves bullet points and list structure
- Removes unsafe elements and inline styles
- Converts images to base64 inline format
- Modern Tailwind CSS interface
- Progress tracking during upload
- Copy HTML to clipboard functionality
- Live preview of converted HTML

## Usage

1. Upload a .docx or .doc file (max 10MB)
2. Click "Convert to HTML"
3. View the clean HTML output
4. Copy or preview the results

