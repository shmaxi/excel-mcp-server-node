import ExcelJS from 'exceljs';
import { Tool, CellFormat } from '../types/index.js';

export const formatTools: Tool[] = [
  {
    name: 'format_range',
    description: 'Apply formatting to a range of cells',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheet_name: {
          type: 'string',
          description: 'Name of the worksheet',
        },
        range: {
          type: 'string',
          description: 'Cell range to format (e.g., "A1:D10")',
        },
        format: {
          type: 'object',
          description: 'Formatting options',
          properties: {
            font: {
              type: 'object',
              properties: {
                name: { type: 'string' },
                size: { type: 'number' },
                bold: { type: 'boolean' },
                italic: { type: 'boolean' },
                underline: { type: 'boolean' },
                color: { type: 'string' },
              },
            },
            fill: {
              type: 'object',
              properties: {
                type: { type: 'string' },
                pattern: { type: 'string' },
                color: { type: 'string' },
                bgColor: { type: 'string' },
              },
            },
            alignment: {
              type: 'object',
              properties: {
                horizontal: { type: 'string' },
                vertical: { type: 'string' },
                wrapText: { type: 'boolean' },
              },
            },
            border: {
              type: 'object',
            },
            numFmt: { type: 'string' },
          },
        },
      },
      required: ['file_path', 'sheet_name', 'range', 'format'],
    },
    handler: async (args: {
      file_path: string;
      sheet_name: string;
      range: string;
      format: CellFormat;
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
        
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startCol; col <= endCol; col++) {
            const cell = worksheet.getCell(row, col);
            
            if (args.format.font) {
              cell.font = {
                ...args.format.font,
                color: args.format.font.color ? { argb: args.format.font.color } : undefined,
              };
            }
            
            if (args.format.fill) {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: args.format.fill.color || 'FFFFFF' },
              };
            }
            
            if (args.format.alignment) {
              cell.alignment = args.format.alignment;
            }
            
            if (args.format.numFmt) {
              cell.numFmt = args.format.numFmt;
            }
          }
        }
      }
      
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Formatting applied to range ${args.range}`,
      };
    },
  },
  
  {
    name: 'merge_cells',
    description: 'Merge a range of cells',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheet_name: {
          type: 'string',
          description: 'Name of the worksheet',
        },
        range: {
          type: 'string',
          description: 'Cell range to merge (e.g., "A1:D1")',
        },
      },
      required: ['file_path', 'sheet_name', 'range'],
    },
    handler: async (args: {
      file_path: string;
      sheet_name: string;
      range: string;
    }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const worksheet = workbook.getWorksheet(args.sheet_name);
      if (!worksheet) {
        throw new Error(`Worksheet ${args.sheet_name} not found`);
      }
      
      worksheet.mergeCells(args.range);
      
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Cells merged: ${args.range}`,
      };
    },
  },
  
  {
    name: 'unmerge_cells',
    description: 'Unmerge previously merged cells',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheet_name: {
          type: 'string',
          description: 'Name of the worksheet',
        },
        range: {
          type: 'string',
          description: 'Cell range to unmerge (e.g., "A1:D1")',
        },
      },
      required: ['file_path', 'sheet_name', 'range'],
    },
    handler: async (args: {
      file_path: string;
      sheet_name: string;
      range: string;
    }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const worksheet = workbook.getWorksheet(args.sheet_name);
      if (!worksheet) {
        throw new Error(`Worksheet ${args.sheet_name} not found`);
      }
      
      worksheet.unMergeCells(args.range);
      
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Cells unmerged: ${args.range}`,
      };
    },
  },
  
  {
    name: 'apply_formula',
    description: 'Apply a formula to a cell',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheet_name: {
          type: 'string',
          description: 'Name of the worksheet',
        },
        cell: {
          type: 'string',
          description: 'Cell reference (e.g., "A1")',
        },
        formula: {
          type: 'string',
          description: 'Excel formula (e.g., "=SUM(A1:A10)")',
        },
      },
      required: ['file_path', 'sheet_name', 'cell', 'formula'],
    },
    handler: async (args: {
      file_path: string;
      sheet_name: string;
      cell: string;
      formula: string;
    }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const worksheet = workbook.getWorksheet(args.sheet_name);
      if (!worksheet) {
        throw new Error(`Worksheet ${args.sheet_name} not found`);
      }
      
      const targetCell = worksheet.getCell(args.cell);
      targetCell.value = { formula: args.formula };
      
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Formula applied to cell ${args.cell}`,
        formula: args.formula,
      };
    },
  },
];