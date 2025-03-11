const XLSX = require('xlsx');
const logger = require('./logger');
const stringSimilarity = require('string-similarity');

class HeaderAnalyzer {
    /**
     * @param {object} worksheet - The worksheet object from XLSX.
     * @param {object} range - The decoded range of the worksheet.
     * @param {Array} customPatterns - Custom regex patterns for header detection.
     * @param {object} detectionConfig - Additional configuration for header detection.
     */
    constructor(worksheet, range, customPatterns = [], detectionConfig = {}) {
        this.worksheet = worksheet;
        this.range = range;
        this.customPatterns = customPatterns;
        this.detectionConfig = detectionConfig;
    }

    /**
     * Returns an array of header objects for the given row.
     * Each header object contains:
     * - row: row index
     * - column: column index
     * - value: header text
     * - level: header level (default 0, can be modified if merged cells indicate hierarchy)
     */
    getHeaders(row) {
        const headers = [];
        for (let col = 0; col <= this.range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = this.worksheet[cellAddress];
            let value = cell ? cell.v : null;
            if (value && typeof value === 'string' && value.trim().length > 0) {
                // Check custom patterns
                let isHeader = false;
                for (const pattern of this.customPatterns) {
                    if (new RegExp(pattern, 'i').test(value)) {
                        isHeader = true;
                        break;
                    }
                }
                // Basic check: contains letters
                if (!isHeader) {
                    isHeader = /[A-Za-z]/.test(value);
                }
                if (isHeader) {
                    headers.push({
                        row: row,
                        column: col,
                        value: value.trim(),
                        level: 0 // Default level; can be updated based on merged cells
                    });
                }
            }
        }
        // Handle merged cells if any (update header levels)
        this.adjustForMergedCells(headers);
        return headers;
    }

    /**
     * Adjust header levels based on merged cells.
     * For simplicity, if a header cell is part of a merged range, set its level accordingly.
     */
    adjustForMergedCells(headers) {
        const merges = this.worksheet['!merges'] || [];
        merges.forEach(merge => {
            headers.forEach(header => {
                // If header cell is within a merged range, update level based on rows spanned
                if (header.column >= merge.s.c && header.column <= merge.e.c &&
                    header.row === merge.s.r) {
                    header.level = merge.e.r - merge.s.r; // Simple heuristic: level equals number of rows spanned
                }
            });
        });
    }
}

module.exports = HeaderAnalyzer;
