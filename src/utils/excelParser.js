const XLSX = require('xlsx');
const logger = require('./logger');
const config = require('../config');
const stringSimilarity = require('string-similarity');
const { v4: uuidv4 } = require('uuid');
const HeaderAnalyzer = require('./headerAnalyzer');

class ExcelParser {
    constructor(worksheet) {
        this.worksheet = worksheet;
        this.range = XLSX.utils.decode_range(worksheet['!ref']);
        this.headerCache = new Map();
        this.tableStructures = [];
        this.dataCache = new Map();
        this.config = config.get('excel');
        this.errors = [];
        this.customPatterns = this.config.customPatterns || [];
        this.performanceMetrics = {
            startTime: Date.now(),
            tablesProcessed: 0,
            rowsProcessed: 0,
            errorsEncountered: 0
        };

        // Initialize header analyzer for complex header extraction
        this.headerAnalyzer = new HeaderAnalyzer(worksheet, this.range, this.customPatterns, this.config.headerDetection);
    }

    /**
     * Searches the worksheet for a cell whose value matches any of the provided header keywords.
     * When found, it returns the value of the cell immediately to its right.
     *
     * @param {string|string[]} keywords - A keyword or array of keywords to search for.
     * @returns {*} The cell value adjacent to the matched header, or null if no match is found.
     */
    getValueByHeader(keywords) {
        if (!Array.isArray(keywords)) {
            keywords = [keywords];
        }
        // Limit search to the first 20 rows for performance.
        const maxRows = Math.min(20, this.range.e.r + 1);
        for (let row = 0; row < maxRows; row++) {
            for (let col = 0; col <= this.range.e.c; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                const cell = this.worksheet[cellAddress];
                if (cell && typeof cell.v === 'string') {
                    const cellValue = cell.v.trim().toLowerCase();
                    for (const keyword of keywords) {
                        if (cellValue.includes(keyword.toLowerCase())) {
                            // If found, return the value in the adjacent cell (to the right)
                            const adjacentAddress = XLSX.utils.encode_cell({ r: row, c: col + 1 });
                            const adjacentCell = this.worksheet[adjacentAddress];
                            if (adjacentCell) {
                                return adjacentCell.v;
                            }
                        }
                    }
                }
            }
        }
        return null;
    }

    /**
     * Find table structures in the worksheet.
     * This method detects basic tables, applies hierarchical analysis,
     * then groups and merges related tables.
     */
    findTableStructures() {
        try {
            const basicTables = this.findBasicTables();
            const hierarchicalTables = this.analyzeTableHierarchy(basicTables);
            const mergedTables = this.processAndMergeTables(hierarchicalTables);
            this.tableStructures = mergedTables;
            this.performanceMetrics.tablesProcessed = mergedTables.length;
            logger.info('Table processing completed', { metrics: this.performanceMetrics });
            return mergedTables;
        } catch (error) {
            this.handleError('table-detection', error);
            return [];
        }
    }

    /**
     * Scan through rows to detect basic tables.
     * A table starts when a header row is detected and ends when there are two or more consecutive empty rows.
     */
    findBasicTables() {
        const tables = [];
        let currentTable = null;
        let tableId = 0;
        try {
            for (let row = 0; row <= this.range.e.r; row++) {
                const rowAnalysis = this.analyzeRow(row);
                if (rowAnalysis.isHeaderRow && !currentTable) {
                    currentTable = this.initializeTable(row, tableId++);
                } else if (currentTable) {
                    if (rowAnalysis.isEmpty && rowAnalysis.consecutiveEmptyRows > 1) {
                        this.finalizeTable(currentTable, row - 1);
                        tables.push(currentTable);
                        currentTable = null;
                    } else if (!rowAnalysis.isEmpty) {
                        currentTable.endRow = row;
                        if (rowAnalysis.isDataRow) {
                            currentTable.dataRows.push(this.extractRowData(row, currentTable.headers));
                        }
                    }
                }
            }
            if (currentTable) {
                this.finalizeTable(currentTable, this.range.e.r);
                tables.push(currentTable);
            }
            return tables;
        } catch (error) {
            this.handleError('basic-table-detection', error);
            return [];
        }
    }

    /**
     * Analyze a given row to determine:
     * - Whether it is a header row (based on the number of header-like cells)
     * - Whether it is a data row
     * - Whether it is empty
     */
    analyzeRow(row) {
        let headerCount = 0;
        let dataCount = 0;
        let emptyCount = 0;
        for (let col = 0; col <= this.range.e.c; col++) {
            const value = this.getCellValue(row, col);
            if (value === null || value === undefined || value === "") {
                emptyCount++;
                continue;
            }
            if (this.isPotentialHeader(value)) {
                headerCount++;
            } else if (this.isPotentialData(value)) {
                dataCount++;
            }
        }
        return {
            isHeaderRow: headerCount >= 2, // heuristic: at least two header-like cells
            isDataRow: dataCount > 0,
            isEmpty: emptyCount === (this.range.e.c + 1),
            consecutiveEmptyRows: this.countConsecutiveEmptyRows(row),
            headerCount,
            dataCount
        };
    }

    // Helper: Get the cell value from a given row and column.
    getCellValue(row, col) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = this.worksheet[cellAddress];
        return cell ? cell.v : null;
    }

    // Helper: Count consecutive empty rows starting at the given row.
    countConsecutiveEmptyRows(currentRow) {
        let count = 0;
        for (let row = currentRow + 1; row <= this.range.e.r; row++) {
            let isEmpty = true;
            for (let col = 0; col <= this.range.e.c; col++) {
                const value = this.getCellValue(row, col);
                if (value !== null && value !== undefined && value !== "") {
                    isEmpty = false;
                    break;
                }
            }
            if (isEmpty) {
                count++;
            } else {
                break;
            }
        }
        return count;
    }

    // Determine if a cell value is likely a header.
    isPotentialHeader(value) {
        if (typeof value !== 'string') return false;
        // Check against custom header patterns
        for (const pattern of this.customPatterns) {
            if (new RegExp(pattern, 'i').test(value)) {
                return true;
            }
        }
        return value.trim().length > 0 && /[A-Za-z]/.test(value);
    }

    // Simple heuristic: a non-header cell with content is considered data.
    isPotentialData(value) {
        if (typeof value === 'string' && value.trim().length > 0) return true;
        if (typeof value === 'number' || typeof value === 'boolean') return true;
        return false;
    }

    // Initialize a new table structure.
    initializeTable(row, id) {
        return {
            id: `table_${id}`,
            startRow: row,
            endRow: row,
            headers: this.headerAnalyzer.getHeaders(row),
            dataRows: [],
            children: [],
            level: 0,
            parent: null,
            metadata: {}
        };
    }

    // Finalize the table by setting its end row and extracting metadata.
    finalizeTable(table, endRow) {
        table.endRow = endRow;
        table.rowCount = table.dataRows.length;
        table.metadata = this.extractTableMetadata(table);
    }

    // Extract metadata for a table (indentation, pattern type, content type).
    extractTableMetadata(table) {
        return {
            indentation: this.getTableIndentation(table),
            patternType: this.identifyTablePattern(table),
            contentType: this.identifyContentType(table)
        };
    }

    // Dummy implementation: return the starting column of the first header as the table indentation.
    getTableIndentation(table) {
        if (table.headers.length === 0) return 0;
        return table.headers[0].column;
    }

    // Dummy implementation: for now, return 'default' as the table pattern.
    identifyTablePattern(table) {
        return 'default';
    }

    // Dummy implementation: determine content type based on the first data row.
    identifyContentType(table) {
        if (table.dataRows.length === 0) return 'unknown';
        const firstRow = table.dataRows[0];
        if (Object.values(firstRow).some(val => typeof val === 'number')) return 'numeric';
        return 'text';
    }

    // Extract data from a given row using the table's headers.
    extractRowData(row, headers) {
        const rowData = {};
        headers.forEach(header => {
            const value = this.getCellValue(row, header.column);
            if (value !== null && value !== undefined) {
                rowData[header.value] = this.formatCellValue(value);
            }
        });
        return rowData;
    }

    // Format a cell's value (placeholder – can be extended for date formatting, etc.).
    formatCellValue(value) {
        return value;
    }

    /**
     * Analyze hierarchical relationships among tables.
     * This method assigns parent/child relationships based on row proximity and header indentation.
     */
    analyzeTableHierarchy(tables) {
        const hierarchicalTables = [];
        const processedIndices = new Set();

        try {
            for (let i = 0; i < tables.length; i++) {
                if (processedIndices.has(i)) continue;
                const currentTable = tables[i];
                currentTable.children = [];
                currentTable.level = 0;
                currentTable.parent = null;

                for (let j = i + 1; j < tables.length; j++) {
                    if (processedIndices.has(j)) continue;
                    const potentialChild = tables[j];
                    if (this.isChildTable(currentTable, potentialChild)) {
                        potentialChild.level = currentTable.level + 1;
                        potentialChild.parent = currentTable;
                        currentTable.children.push(potentialChild);
                        processedIndices.add(j);
                    }
                }
                hierarchicalTables.push(currentTable);
                processedIndices.add(i);
            }
            this.analyzeTableRelationships(hierarchicalTables);
            this.validateTableHierarchy(hierarchicalTables);
            return hierarchicalTables;
        } catch (error) {
            this.handleError('hierarchy-analysis', error);
            return tables;
        }
    }

    // Heuristic: determine if potentialChild is a child of parentTable.
    isChildTable(parentTable, potentialChild) {
        // Example: if the potential child immediately follows and is more indented.
        if (potentialChild.startRow === parentTable.endRow + 1) {
            const parentIndent = parentTable.headers.length > 0 ? parentTable.headers[0].column : 0;
            const childIndent = potentialChild.headers.length > 0 ? potentialChild.headers[0].column : 0;
            return childIndent > parentIndent;
        }
        return false;
    }

    // Label table relationships (parent, child, or independent).
    analyzeTableRelationships(tables) {
        tables.forEach(table => {
            if (table.children && table.children.length > 0) {
                table.metadata.relationshipType = 'parent';
            } else if (table.parent) {
                table.metadata.relationshipType = 'child';
            } else {
                table.metadata.relationshipType = 'independent';
            }
        });
    }

    // Validate hierarchy to ensure no circular references exist.
    validateTableHierarchy(tables) {
        const visited = new Set();
        const validateTable = (table) => {
            if (visited.has(table.id)) {
                throw new Error(`Circular reference detected in table hierarchy: ${table.id}`);
            }
            visited.add(table.id);
            table.children.forEach(validateTable);
        };
        tables.forEach(validateTable);
    }

    /**
     * Process and merge related tables.
     * This groups tables (based on header similarity and row proximity) and then merges each group.
     */
    processAndMergeTables(tables) {
        try {
            const tableGroups = this.groupRelatedTables(tables);
            const mergedTables = tableGroups.map(group => this.mergeTables(group));
            mergedTables.forEach(table => this.processTableData(table));
            return mergedTables;
        } catch (error) {
            this.handleError('table-merging', error);
            return tables;
        }
    }

    // Group tables into clusters based on header similarity and proximity.
    groupRelatedTables(tables) {
        const groups = [];
        const used = new Set();

        for (let i = 0; i < tables.length; i++) {
            if (used.has(i)) continue;
            const group = [tables[i]];
            used.add(i);
            for (let j = i + 1; j < tables.length; j++) {
                if (used.has(j)) continue;
                if (this.areTablesRelated(tables[i], tables[j])) {
                    group.push(tables[j]);
                    used.add(j);
                }
            }
            groups.push(group);
        }
        return groups;
    }

    // Heuristic: two tables are related if they are close in rows and have similar headers.
    areTablesRelated(table1, table2) {
        if (Math.abs(table2.startRow - table1.endRow) > 5) return false; // too far apart
        let matchCount = 0;
        table1.headers.forEach(h1 => {
            table2.headers.forEach(h2 => {
                if (this.areHeadersSimilar(h1, h2)) matchCount++;
            });
        });
        const threshold = Math.min(table1.headers.length, table2.headers.length) * 0.5;
        return matchCount >= threshold;
    }

    // Merge a group of related tables into a single table.
    mergeTables(group) {
        if (group.length === 1) return group[0];
        const baseTable = group[0];
        for (let i = 1; i < group.length; i++) {
            const table = group[i];
            baseTable.headers = this.mergeHeaders(baseTable.headers, table.headers);
            baseTable.dataRows = baseTable.dataRows.concat(table.dataRows);
            baseTable.endRow = Math.max(baseTable.endRow, table.endRow);
        }
        return baseTable;
    }

    // Merge two header arrays by uniting unique headers based on similarity.
    mergeHeaders(headers1, headers2) {
        const merged = [...headers1];
        headers2.forEach(h2 => {
            const exists = merged.some(h1 => this.areHeadersSimilar(h1, h2));
            if (!exists) {
                merged.push(h2);
            }
        });
        return merged;
    }

    // Process table data after merging (e.g., remove duplicate rows).
    processTableData(table) {
        const uniqueRows = [];
        table.dataRows.forEach(row => {
            if (!this.findMatchingRow(uniqueRows, row)) {
                uniqueRows.push(row);
            }
        });
        table.dataRows = uniqueRows;
    }

    // Find a matching row in an array of rows based on a similarity threshold.
    findMatchingRow(rows, newRow) {
        return rows.find(row => {
            const keys = Object.keys(newRow);
            let matchScore = 0;
            keys.forEach(key => {
                if (row[key] === newRow[key]) matchScore++;
            });
            return matchScore / keys.length > 0.7;
        });
    }

    /**
     * Advanced pattern-based data grouping.
     * This method scans through tables and headers to extract data groups and patterns.
     */
    findRelatedDataGroups() {
        const groups = new Map();
        const tables = this.findTableStructures();
        try {
            tables.forEach(table => {
                table.headers.forEach(header => {
                    const groupType = this.identifyDataGroup(header.value);
                    if (groupType) {
                        if (!groups.has(groupType)) {
                            groups.set(groupType, {
                                type: groupType,
                                locations: [],
                                relatedData: new Map(),
                                patterns: []
                            });
                        }
                        const group = groups.get(groupType);
                        group.locations.push({
                            row: table.startRow,
                            col: header.column,
                            table: table
                        });
                        const pattern = this.extractDataPattern(table, header);
                        if (pattern) {
                            group.patterns.push(pattern);
                        }
                    }
                });
            });
            groups.forEach((group) => {
                group.locations.forEach(location => {
                    const relatedData = this.findRelatedData(location.row, location.col, location.table);
                    this.mergeRelatedData(group.relatedData, relatedData);
                });
            });
            return groups;
        } catch (error) {
            this.handleError('data-grouping', error);
            return new Map();
        }
    }

    // Dummy implementation: determine group type from a header value.
    identifyDataGroup(headerValue) {
        const lower = headerValue.toLowerCase();
        if (lower.includes('product')) return 'product';
        if (lower.includes('date')) return 'date';
        return null;
    }

    // Extract a data pattern for a given header in a table.
    extractDataPattern(table, header) {
        const values = table.dataRows.map(row => row[header.value]).filter(Boolean);
        if (values.length === 0) return null;
        return {
            type: this.identifyValueType(values[0]),
            format: this.detectFormat(values),
            isUnique: new Set(values).size === values.length,
            hasNulls: values.some(v => v === null || v === undefined),
            examples: values.slice(0, 3)
        };
    }

    // Determine the value type of a sample (number, boolean, date, code, text).
    identifyValueType(value) {
        if (typeof value === 'number') return 'number';
        if (typeof value === 'boolean') return 'boolean';
        if (typeof value === 'string') {
            if (/^\d{4}-\d{2}-\d{2}/.test(value)) return 'date';
            if (/^\d{2}\/\d{2}\/\d{4}/.test(value)) return 'date';
            if (/^[A-Z0-9-]+$/.test(value)) return 'code';
        }
        return 'text';
    }

    // Detect format information based on sample values.
    detectFormat(values) {
        const sample = values[0].toString();
        if (sample.includes('/') || sample.includes('-')) {
            const dateFormat = this.detectDateFormat(sample);
            if (dateFormat) return { type: 'date', format: dateFormat };
        }
        if (!isNaN(sample)) {
            return { type: 'number', format: this.detectNumberFormat(sample) };
        }
        if (/^[A-Z0-9-]+$/.test(sample)) {
            return { type: 'code', format: this.detectCodePattern(sample) };
        }
        return { type: 'text', format: null };
    }

    // Dummy date format detector.
    detectDateFormat(sample) {
        if (/^\d{4}-\d{2}-\d{2}/.test(sample)) return 'YYYY-MM-DD';
        if (/^\d{2}\/\d{2}\/\d{4}/.test(sample)) return 'MM/DD/YYYY';
        return null;
    }

    // Dummy number format detector.
    detectNumberFormat(sample) {
        return 'standard';
    }

    // Dummy code pattern detector.
    detectCodePattern(sample) {
        return 'alphanumeric';
    }

    // Merge related data from source into a target map.
    mergeRelatedData(target, source) {
        source.forEach((value, key) => {
            if (!target.has(key)) {
                target.set(key, []);
            }
            target.get(key).push(value);
        });
    }

    // Determine if two headers are similar using fuzzy string matching.
    areHeadersSimilar(header1, header2) {
        return stringSimilarity.compareTwoStrings(
            header1.value.toLowerCase(),
            header2.value.toLowerCase()
        ) > this.config.headerDetection.fuzzyMatchThreshold;
    }

    // Error handling: logs error details and optionally throws errors.
    handleError(context, error) {
        const errorInfo = {
            id: uuidv4(),
            context,
            message: error.message,
            stack: error.stack,
            timestamp: new Date().toISOString()
        };
        this.errors.push(errorInfo);
        this.performanceMetrics.errorsEncountered++;
        logger.error(`Error in ${context}:`, errorInfo);
        if (this.config.errorHandling && this.config.errorHandling.throwErrors) {
            throw error;
        }
    }

    // Public method: Return product data.
    findProductData() {
        const tables = this.findTableStructures();
        if (tables.length === 0) return [];
        // Assumes that the first table contains product data.
        return tables[0].dataRows;
    }

    // Parse a date from a cell value; returns a Date object or null if invalid.
    parseDate(value) {
        const parsed = new Date(value);
        return isNaN(parsed.getTime()) ? null : parsed;
    }

    // Return the error log.
    getErrors() {
        return [...this.errors];
    }

    // Clear the error log.
    clearErrors() {
        this.errors = [];
    }
}

module.exports = ExcelParser;
