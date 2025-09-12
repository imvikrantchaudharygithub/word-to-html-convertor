"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.convertDocxToHtml = convertDocxToHtml;
exports.isValidWordFile = isValidWordFile;
const mammoth = __importStar(require("mammoth"));
const jsdom_1 = require("jsdom");
const dompurify_1 = __importDefault(require("dompurify"));
/**
 * Converts a Word document buffer to clean HTML with proper heading detection
 * @param buffer - The Word document buffer
 * @returns Promise<string> - Clean HTML string
 */
async function convertDocxToHtml(buffer) {
    try {
        // Step 1: First, let's extract the raw document to understand its structure
        console.log('Step 1: Analyzing document structure...');
        const rawResult = await mammoth.extractRawText({ buffer });
        console.log('Raw text length:', rawResult.value.length);
        // Debug: Extract document info to understand structure
        console.log('Step 1.5: Extracting document info for debugging...');
        // Extract document structure info
        const docInfo = await mammoth.convertToHtml({ buffer }, {
            includeDefaultStyleMap: true,
            includeEmbeddedStyleMap: true
        });
        console.log('Mammoth messages:', docInfo.messages?.map(m => `${m.type}: ${m.message}`));
        // Step 2: Convert with comprehensive style mapping
        console.log('Step 2: Converting with enhanced style mapping...');
        // First, let's test without our style mappings to see if that's the issue
        console.log('Step 2a: Testing basic conversion first...');
        const basicTest = await mammoth.convertToHtml({ buffer });
        const basicUlCount = (basicTest.value.match(/<ul>/g) || []).length;
        const basicLiCount = (basicTest.value.match(/<li>/g) || []).length;
        console.log(`Basic conversion has: ${basicUlCount} lists, ${basicLiCount} items`);
        // Use basic conversion to preserve lists, then enhance with our processing
        const result = basicTest;
        // Log conversion messages for debugging
        if (result.messages.length > 0) {
            console.log('Mammoth conversion messages:');
            result.messages.forEach(msg => console.log(`- ${msg.type}: ${msg.message}`));
        }
        console.log('Step 3: Post-processing HTML...');
        // Step 3: Advanced post-processing to catch missed headings and fix lists
        console.log('Before heading detection - lists check:');
        const beforeUlCount = (result.value.match(/<ul>/g) || []).length;
        const beforeLiCount = (result.value.match(/<li>/g) || []).length;
        console.log(`Before: ${beforeUlCount} lists, ${beforeLiCount} items`);
        let processedHtml = advancedHeadingDetection(result.value);
        console.log('After heading detection - lists check:');
        const afterHeadingUlCount = (processedHtml.match(/<ul>/g) || []).length;
        const afterHeadingLiCount = (processedHtml.match(/<li>/g) || []).length;
        console.log(`After heading detection: ${afterHeadingUlCount} lists, ${afterHeadingLiCount} items`);
        console.log('Step 3.5: Processing bullet points and lists...');
        processedHtml = enhancedListProcessing(processedHtml);
        console.log('Step 3.7: Removing strong tags and improving heading structure...');
        processedHtml = removeStrongTagsAndImproveHeadings(processedHtml);
        console.log('Step 4: Sanitizing HTML...');
        // Step 4: Sanitize and clean the HTML
        const cleanHtml = sanitizeHtml(processedHtml);
        console.log(`Final HTML length: ${cleanHtml.length} characters`);
        return cleanHtml;
    }
    catch (error) {
        console.error('Error converting DOCX to HTML:', error);
        throw new Error('Failed to convert Word document to HTML');
    }
}
/**
 * Advanced heading detection using multiple heuristics
 * This function analyzes text formatting, positioning, and content patterns
 * @param html - Raw HTML from mammoth
 * @returns string - HTML with proper heading tags
 */
function advancedHeadingDetection(html) {
    console.log('Running advanced heading detection...');
    // Create a DOM environment to manipulate HTML
    const window = new jsdom_1.JSDOM(html).window;
    const document = window.document;
    // Find all paragraphs that might be headings
    const paragraphs = document.querySelectorAll('p');
    let headingCount = 0;
    paragraphs.forEach((p, index) => {
        const text = p.textContent?.trim() || '';
        // Skip empty paragraphs
        if (!text)
            return;
        // Analyze if this paragraph should be a heading
        const headingInfo = analyzeHeadingCandidate(p, text, index, paragraphs.length);
        if (headingInfo.isHeading) {
            // Create new heading element
            const heading = document.createElement(headingInfo.level);
            // Preserve any formatting within the heading
            heading.innerHTML = p.innerHTML;
            // Replace paragraph with heading
            p.parentNode?.replaceChild(heading, p);
            headingCount++;
            console.log(`Converted to ${headingInfo.level}: "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}"`);
        }
    });
    console.log(`Converted ${headingCount} paragraphs to headings`);
    return document.body.innerHTML;
}
/**
 * Analyzes a paragraph to determine if it should be converted to a heading
 * Uses conservative heuristics to avoid converting list items and regular content
 */
function analyzeHeadingCandidate(p, text, index, totalParagraphs) {
    let score = 0;
    let suggestedLevel = 'h2';
    // EXCLUSION RULES FIRST - These prevent false positives
    // Exclude if it's clearly a list item or bullet point (expanded patterns)
    if (/^[•·▪▫◦‣⁃○●■□▲►]\s/.test(text) || /^[-*+→]\s/.test(text)) {
        return { isHeading: false, level: 'h2' };
    }
    // Exclude if it starts with a number followed by a period and space (likely enumerated list)
    if (/^\d+[.)]\s/.test(text) || /^[a-zA-Z][.)]\s/.test(text)) {
        return { isHeading: false, level: 'h2' };
    }
    // Exclude if it starts with Roman numerals (common in lists)
    if (/^[ivxlcdm]+[.)]\s/i.test(text)) {
        return { isHeading: false, level: 'h2' };
    }
    // Exclude if it's part of a nested structure (contains multiple indented items)
    if (text.includes('○') || text.includes('■') || text.includes('▪')) {
        return { isHeading: false, level: 'h2' };
    }
    // Exclude very long paragraphs (likely body text)
    if (text.length > 150) {
        return { isHeading: false, level: 'h2' };
    }
    // Exclude if it ends with common sentence punctuation
    if (text.endsWith('.') || text.endsWith(',') || text.endsWith(';') || text.endsWith(':')) {
        return { isHeading: false, level: 'h2' };
    }
    // POSITIVE INDICATORS - Only for clear heading patterns
    // Strong indicator: Entirely bold text that's reasonably short
    const strongElements = p.querySelectorAll('strong, b');
    const isEntirelyBold = strongElements.length > 0 &&
        strongElements[0] &&
        strongElements[0].textContent?.trim() === text;
    if (isEntirelyBold) {
        if (text.length <= 150) {
            score += 70; // Strong indicator for bold text (increased threshold)
        }
        else if (text.length <= 250) {
            score += 50; // Medium indicator for longer bold text
        }
    }
    // Major section headings (common patterns)
    const majorHeadingPatterns = [
        /^(overview|introduction|conclusion|summary|background|methodology|results|discussion|references|abstract)/i,
        /^(chapter|section|part|appendix)\s+\d+/i,
        /^[A-Z][a-z]+(\s+[A-Z][a-z]+)*:?\s*$/, // Title Case patterns
        /^[A-Z\s]+$/ // ALL CAPS (but not too long)
    ];
    if (majorHeadingPatterns.some(pattern => pattern.test(text)) && text.length <= 60) {
        score += 40;
        suggestedLevel = 'h1';
    }
    // Question patterns (FAQ style)
    if (text.endsWith('?') && text.length <= 100) {
        score += 35;
        suggestedLevel = 'h2';
    }
    // Numbered section headings (but not list items)
    if (/^\d+\.?\s+[A-Z]/.test(text) && text.length <= 80) {
        score += 30;
        const match = text.match(/^(\d+)/);
        if (match && match[1]) {
            const num = parseInt(match[1]);
            suggestedLevel = num <= 3 ? 'h2' : 'h3';
        }
    }
    // Position bonus for early document headings
    const relativePosition = index / totalParagraphs;
    if (relativePosition < 0.05 && text.length <= 100) {
        score += 20;
        suggestedLevel = 'h1';
    }
    // Very conservative threshold - only convert obvious headings
    const isHeading = score >= 70;
    if (isHeading) {
        console.log(`Converting to heading: "${text.substring(0, 50)}..." - Score: ${score}`);
    }
    return {
        isHeading,
        level: suggestedLevel
    };
}
/**
 * Sanitizes HTML by removing inline styles, classes, and unsafe elements
 * @param html - Raw HTML string
 * @returns string - Sanitized HTML string
 */
function sanitizeHtml(html) {
    // Create a DOM environment for DOMPurify
    const window = new jsdom_1.JSDOM('').window;
    const purify = (0, dompurify_1.default)(window);
    // Configure DOMPurify to remove unsafe elements and attributes
    const cleanHtml = purify.sanitize(html, {
        // Allow these HTML tags
        ALLOWED_TAGS: [
            'p', 'br', 'em', 'i', 'u', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
            'ul', 'ol', 'li', 'a', 'img', 'table', 'thead', 'tbody', 'tr', 'td', 'th',
            'blockquote', 'div', 'span', 'sup', 'sub'
        ],
        // Allow these attributes
        ALLOWED_ATTR: ['href', 'src', 'alt', 'title'],
        // Remove all unsafe tags completely
        FORBID_TAGS: ['script', 'iframe', 'object', 'embed', 'form', 'input', 'button', 'select', 'textarea'],
        // Remove these attributes (including all style and class attributes)
        FORBID_ATTR: ['style', 'class', 'id', 'onclick', 'onload', 'onerror', 'onmouseover'],
        // Keep content when removing forbidden tags
        KEEP_CONTENT: true,
        // Remove empty attributes
        ALLOW_ARIA_ATTR: false,
        ALLOW_DATA_ATTR: false
    });
    // Additional cleanup: remove any remaining inline styles and classes
    const finalHtml = cleanHtml
        .replace(/\s+style\s*=\s*["'][^"']*["']/gi, '') // Remove any remaining style attributes
        .replace(/\s+class\s*=\s*["'][^"']*["']/gi, '') // Remove any remaining class attributes
        .replace(/\s+id\s*=\s*["'][^"']*["']/gi, '') // Remove any remaining id attributes
        .replace(/\s{2,}/g, ' ') // Replace multiple spaces with single space
        .trim();
    return finalHtml;
}
/**
 * Validates if the file is a supported Word document format
 * @param filename - The name of the uploaded file
 * @param mimetype - The MIME type of the uploaded file
 * @returns boolean - True if file is supported
 */
/**
 * Enhanced list processing to preserve existing lists and clean up any formatting issues
 * @param html - HTML string with existing lists
 * @returns string - HTML with preserved and cleaned lists
 */
function enhancedListProcessing(html) {
    console.log('Running enhanced list processing...');
    const window = new jsdom_1.JSDOM(html).window;
    const document = window.document;
    // Count existing lists
    const existingUlCount = document.querySelectorAll('ul').length;
    const existingOlCount = document.querySelectorAll('ol').length;
    const existingLiCount = document.querySelectorAll('li').length;
    console.log(`Found ${existingUlCount} unordered lists, ${existingOlCount} ordered lists, ${existingLiCount} list items`);
    // Clean up any formatting issues in existing lists
    const allListItems = document.querySelectorAll('li');
    let cleanedItems = 0;
    allListItems.forEach(li => {
        // Remove any strong tags from list items
        const strongElements = li.querySelectorAll('strong');
        strongElements.forEach(strong => {
            const textNode = document.createTextNode(strong.textContent || '');
            if (strong.parentNode) {
                strong.parentNode.replaceChild(textNode, strong);
                cleanedItems++;
            }
        });
    });
    console.log(`Cleaned formatting in ${cleanedItems} list items`);
    console.log(`Final count: ${existingUlCount + existingOlCount} lists with ${existingLiCount} items`);
    return document.body.innerHTML;
}
/**
 * Determines if a text string represents a list item
 */
function isListItem(text) {
    if (!text || text.length < 2)
        return false;
    // Bullet point patterns
    const bulletPatterns = [
        /^[•·▪▫◦‣⁃○●■□▲►]\s/, // Various bullet symbols
        /^[-*+→]\s/, // Dash, asterisk, plus, arrow
    ];
    // Numbered list patterns
    const numberedPatterns = [
        /^\d+[.)]\s/, // 1. or 1) 
        /^[a-zA-Z][.)]\s/, // a. or a) or A. or A)
        /^[ivxlcdm]+[.)]\s/i, // Roman numerals
    ];
    return bulletPatterns.some(pattern => pattern.test(text)) ||
        numberedPatterns.some(pattern => pattern.test(text));
}
/**
 * Removes strong tags and improves heading structure based on document analysis
 * @param html - HTML string with strong tags and potential headings
 * @returns string - HTML with improved heading structure and no strong tags
 */
function removeStrongTagsAndImproveHeadings(html) {
    console.log('Removing strong tags and improving heading structure...');
    const window = new jsdom_1.JSDOM(html).window;
    const document = window.document;
    // Find all paragraphs with strong tags that should be headings
    const paragraphs = Array.from(document.querySelectorAll('p'));
    let strongTagsRemoved = 0;
    let headingsImproved = 0;
    paragraphs.forEach(p => {
        const strongElements = Array.from(p.querySelectorAll('strong'));
        const text = p.textContent?.trim() || '';
        // Check if this paragraph contains only a strong element (potential heading)
        const hasOnlyStrong = strongElements.length === 1 &&
            strongElements[0] &&
            strongElements[0].textContent?.trim() === text;
        if (hasOnlyStrong && shouldBeHeading(text)) {
            // Convert to heading
            const headingLevel = determineHeadingLevel(text);
            const heading = document.createElement(headingLevel);
            // Move the text content without the strong tag
            heading.textContent = text;
            // Replace paragraph with heading
            if (p.parentNode) {
                p.parentNode.replaceChild(heading, p);
                headingsImproved++;
            }
        }
        else {
            // Remove all strong tags from this paragraph but keep the text
            strongElements.forEach(strong => {
                const textNode = document.createTextNode(strong.textContent || '');
                if (strong.parentNode) {
                    strong.parentNode.replaceChild(textNode, strong);
                    strongTagsRemoved++;
                }
            });
        }
    });
    // Also remove any remaining strong tags from other elements
    const allStrongElements = Array.from(document.querySelectorAll('strong'));
    allStrongElements.forEach(strong => {
        const textNode = document.createTextNode(strong.textContent || '');
        if (strong.parentNode) {
            strong.parentNode.replaceChild(textNode, strong);
            strongTagsRemoved++;
        }
    });
    console.log(`Removed ${strongTagsRemoved} strong tags and improved ${headingsImproved} headings`);
    return document.body.innerHTML;
}
/**
 * Determines if text should be converted to a heading based on content analysis
 */
function shouldBeHeading(text) {
    if (!text || text.length < 3)
        return false;
    // Specific patterns from the analyzed document
    const headingPatterns = [
        // Main title pattern
        /^[A-Z][^:]*:\s*[A-Z]/, // "Title: Subtitle" format
        // Section headers
        /^(Overview|Treatment|Prognosis|Screening|Prevention|Recovery|Frequently Asked Questions|Next Steps)/i,
        // Medical terms that are typically headings
        /^(Surgery|Radiation Therapy|Proton Therapy|Medical Treatment|Causes|Risk Factors|Symptoms|Diagnosis|Staging)/i,
        // Question format (FAQ)
        /^(What|How|Why|When|Is|Will|Can|Are)\s+.*\?$/,
        // "Types and Sites" style headings
        /^[A-Z][a-z]+(\s+(and|or)\s+[A-Z][a-z]+)*$/,
        // Avoid converting if it ends with common sentence punctuation (except questions)
        /[.,:;]$/
    ];
    // Check for heading patterns (exclude the last one which is exclusion)
    const isHeadingPattern = headingPatterns.slice(0, -1).some(pattern => pattern.test(text));
    const lastPattern = headingPatterns[headingPatterns.length - 1];
    const endsWithPunctuation = lastPattern ? lastPattern.test(text) : false;
    // Should be heading if it matches patterns and doesn't end with sentence punctuation
    return isHeadingPattern && !endsWithPunctuation;
}
/**
 * Determines the appropriate heading level based on text content
 */
function determineHeadingLevel(text) {
    // Main document title (long descriptive titles)
    if (text.includes(':') && text.length > 50) {
        return 'h1';
    }
    // Major sections
    if (/^(Overview|Treatment|Prognosis|Screening|Prevention|Recovery|Frequently Asked Questions|Next Steps)/i.test(text)) {
        return 'h2';
    }
    // Treatment subsections
    if (/^(Surgery|Radiation Therapy|Proton Therapy|Medical Treatment|Causes|Risk Factors|Symptoms|Diagnosis|Staging)/i.test(text)) {
        return 'h3';
    }
    // Questions and shorter sections
    if (text.endsWith('?') || text.length < 50) {
        return 'h4';
    }
    // Default to h3 for other cases
    return 'h3';
}
function isValidWordFile(filename, mimetype) {
    const allowedExtensions = ['.docx', '.doc'];
    const allowedMimeTypes = [
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
        'application/msword', // .doc
        'application/vnd.ms-word' // alternative .doc mime type
    ];
    // Check file extension
    const hasValidExtension = allowedExtensions.some(ext => filename.toLowerCase().endsWith(ext));
    // Check MIME type
    const hasValidMimeType = allowedMimeTypes.includes(mimetype);
    return hasValidExtension && hasValidMimeType;
}
//# sourceMappingURL=converter.js.map