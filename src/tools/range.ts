import ExcelJS from 'exceljs';
import { Tool } from '../types/index.js';

export const rangeTools: Tool[] = [
  {
    name: 'copy_range',
    description: 'Copy a range of cells to another location',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        source_sheet: {
          type: 'string',
          description: 'Source worksheet name',
        },
        source_range: {
          type: 'string',
          description: 'Source range (e.g., "A1:C10")',
        },
        target_sheet: {
          type: 'string',
          description: 'Target worksheet name',
        },
        target_cell: {
          type: 'string',
          description: 'Target starting cell (e.g., "E1")',
        },
      },
      required: ['file_path', 'source_sheet', 'source_range', 'target_sheet', 'target_cell'],
    },
    handler: async (args: {
      file_path: string;
      source_sheet: string;
      source_range: string;
      target_sheet: string;
      target_cell: string;
    }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const sourceWorksheet = workbook.getWorksheet(args.source_sheet);
      if (!sourceWorksheet) {
        throw new Error(`Source worksheet ${args.source_sheet} not found`);
      }
      
      const targetWorksheet = workbook.getWorksheet(args.target_sheet);
      if (!targetWorksheet) {
        throw new Error(`Target worksheet ${args.target_sheet} not found`);
      }
      
      const [sourceStart, sourceEnd] = args.source_range.split(':');
      const sourceStartMatch = sourceStart.match(/([A-Z]+)(\d+)/);
      const sourceEndMatch = sourceEnd.match(/([A-Z]+)(\d+)/);
      const targetMatch = args.target_cell.match(/([A-Z]+)(\d+)/);
      
      if (sourceStartMatch && sourceEndMatch && targetMatch) {
        const sourceStartRow = parseInt(sourceStartMatch[2]);
        const sourceEndRow = parseInt(sourceEndMatch[2]);
        const sourceStartCol = sourceStartMatch[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        const sourceEndCol = sourceEndMatch[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        
        const targetStartRow = parseInt(targetMatch[2]);
        const targetStartCol = targetMatch[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        
        for (let row = 0; row <= sourceEndRow - sourceStartRow; row++) {
          for (let col = 0; col <= sourceEndCol - sourceStartCol; col++) {
            const sourceCell = sourceWorksheet.getCell(sourceStartRow + row, sourceStartCol + col);
            const targetCell = targetWorksheet.getCell(targetStartRow + row, targetStartCol + col);
            
            targetCell.value = sourceCell.value;
            targetCell.style = sourceCell.style;
          }
        }
      }
      
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Range ${args.source_range} copied from ${args.source_sheet} to ${args.target_sheet} at ${args.target_cell}`,
      };
    },
  },
  
  {
    name: 'delete_range',
    description: 'Delete a range of cells',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheet_name: {
          type: 'string',
          description: 'Worksheet name',
        },
        range: {
          type: 'string',
          description: 'Range to delete (e.g., "A1:C10")',
        },
        shift: {
          type: 'string',
          description: 'Shift direction after deletion',
          enum: ['up', 'left', 'none'],
          default: 'none',
        },
      },
      required: ['file_path', 'sheet_name', 'range'],
    },
    handler: async (args: {
      file_path: string;
      sheet_name: string;
      range: string;
      shift?: 'up' | 'left' | 'none';
    }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const worksheet = workbook.getWorksheet(args.sheet_name);
      if (!worksheet) {
        throw new Error(`Worksheet ${args.sheet_name} not found`);
      }
      
      const [start, end] = args.range.split(':');
      const startMatch = start.match(/([A-Z]+)(\d+)/);
      const endMatch = end.match(/([A-Z]+)(\d+)/);
      
      if (startMatch && endMatch) {
        const startRow = parseInt(startMatch[2]);
        const endRow = parseInt(endMatch[2]);
        const startCol = startMatch[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        const endCol = endMatch[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        
        // Clear the cells
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startCol; col <= endCol; col++) {
            const cell = worksheet.getCell(row, col);
            cell.value = null;
            cell.style = {};
          }
        }
        
        // Note: ExcelJS doesn't have built-in support for shifting cells
        // This would require more complex implementation
      }
      
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Range ${args.range} cleared`,
        note: args.shift !== 'none' ? 'Cell shifting not implemented in ExcelJS' : undefined,
      };
    },
  },
  
  {
    name: 'validate_excel_range',
    description: 'Validate if a range reference is valid',
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Excel range to validate (e.g., "A1:C10")',
        },
      },
      required: ['range'],
    },
    handler: async (args: { range: string }) => {
      const rangePattern = /^[A-Z]+\d+:[A-Z]+\d+$/;
      const singleCellPattern = /^[A-Z]+\d+$/;
      
      const isValid = rangePattern.test(args.range) || singleCellPattern.test(args.range);
      
      if (isValid && args.range.includes(':')) {
        const [start, end] = args.range.split(':');
        const startMatch = start.match(/([A-Z]+)(\d+)/);
        const endMatch = end.match(/([A-Z]+)(\d+)/);
        
        if (startMatch && endMatch) {
          const startRow = parseInt(startMatch[2]);
          const endRow = parseInt(endMatch[2]);
          const startCol = startMatch[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
          const endCol = endMatch[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
          
          return {
            valid: true,
            range: args.range,
            start: { row: startRow, column: startCol, cell: start },
            end: { row: endRow, column: endCol, cell: end },
            rows: endRow - startRow + 1,
            columns: endCol - startCol + 1,
          };
        }
      } else if (isValid) {
        const match = args.range.match(/([A-Z]+)(\d+)/);
        if (match) {
          const row = parseInt(match[2]);
          const col = match[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
          
          return {
            valid: true,
            range: args.range,
            start: { row, column: col, cell: args.range },
            end: { row, column: col, cell: args.range },
            rows: 1,
            columns: 1,
          };
        }
      }
      
      return {
        valid: false,
        error: 'Invalid range format',
      };
    },
  },
  
  {
    name: 'validate_formula_syntax',
    description: 'Validate Excel formula syntax',
    inputSchema: {
      type: 'object',
      properties: {
        formula: {
          type: 'string',
          description: 'Excel formula to validate',
        },
      },
      required: ['formula'],
    },
    handler: async (args: { formula: string }) => {
      // Basic validation - check if formula starts with =
      if (!args.formula.startsWith('=')) {
        return {
          valid: false,
          error: 'Formula must start with =',
        };
      }
      
      // Check for balanced parentheses
      let parenCount = 0;
      for (const char of args.formula) {
        if (char === '(') parenCount++;
        if (char === ')') parenCount--;
        if (parenCount < 0) {
          return {
            valid: false,
            error: 'Unbalanced parentheses',
          };
        }
      }
      
      if (parenCount !== 0) {
        return {
          valid: false,
          error: 'Unbalanced parentheses',
        };
      }
      
      // Basic function name validation
      const functionPattern = /[A-Z]+\(/g;
      const commonFunctions = ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'IF', 'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH'];
      const matches = args.formula.match(functionPattern);
      
      return {
        valid: true,
        formula: args.formula,
        functions: matches ? matches.map(m => m.slice(0, -1)) : [],
        note: 'Basic syntax validation only. Excel will perform full validation.',
      };
    },
  },
];