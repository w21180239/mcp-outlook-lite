import { Buffer } from 'buffer';
import * as XLSX from 'xlsx';
import { debug } from '../../utils/logger.js';

export function parseExcelContent(contentBytes: string, filename: string, maxSheets = 10, maxRowsPerSheet = 1000) {
  try {
    debug(`Debug: Parsing Excel file: ${filename}`);

    // Decode Base64 to buffer
    const buffer = Buffer.from(contentBytes, 'base64');

    // Parse Excel file
    const workbook = XLSX.read(buffer, { type: 'buffer' });

    const result: Record<string, any> = {
      type: 'excel',
      filename: filename,
      sheets: [],
      summary: {
        totalSheets: workbook.SheetNames.length,
        sheetNames: workbook.SheetNames
      }
    };

    // Process up to maxSheets sheets
    const sheetsToProcess = workbook.SheetNames.slice(0, maxSheets);

    for (const sheetName of sheetsToProcess) {
      debug(`Debug: Processing sheet: ${sheetName}`);

      const worksheet = workbook.Sheets[sheetName];

      // Get sheet range
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
      const totalRows = range.e.r - range.s.r + 1;
      const totalCols = range.e.c - range.s.c + 1;

      // Limit rows to prevent overwhelming output
      const rowsToProcess = Math.min(totalRows, maxRowsPerSheet);

      // Convert to JSON with limited rows
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1, // Use array format instead of object
        range: rowsToProcess < totalRows ? `${(worksheet['!ref'] || 'A1:A1').split(':')[0]}:${XLSX.utils.encode_cell({r: range.s.r + rowsToProcess - 1, c: range.e.c})}` : undefined
      });

      const sheetInfo = {
        name: sheetName,
        dimensions: {
          rows: totalRows,
          columns: totalCols,
          range: worksheet['!ref'] || 'A1:A1'
        },
        data: jsonData,
        truncated: rowsToProcess < totalRows,
        displayedRows: jsonData.length,
        note: rowsToProcess < totalRows ? `Sheet truncated to ${maxRowsPerSheet} rows (total: ${totalRows})` : undefined
      };

      result.sheets.push(sheetInfo);
    }

    if (workbook.SheetNames.length > maxSheets) {
      result.summary.note = `Only first ${maxSheets} sheets displayed (total: ${workbook.SheetNames.length})`;
    }

    debug(`Debug: Successfully parsed Excel file with ${result.sheets.length} sheets`);
    return result;

  } catch (error) {
    debug(`Debug: Excel parsing failed: ${error.message}`);
    return {
      type: 'excel_error',
      error: `Failed to parse Excel file: ${error.message}`,
      note: 'File may be corrupted or in an unsupported Excel format'
    };
  }
}
