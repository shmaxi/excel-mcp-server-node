import ExcelJS from 'exceljs';
import { Tool } from '../types/index.js';

export const dataTools: Tool[] = [
  {
    name: 'write_data_to_excel',
    description: 'Write data to an Excel worksheet',
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
        data: {
          type: 'array',
          description: 'Array of arrays representing rows of data',
          items: {
            type: 'array',
          },
        },
        start_cell: {
          type: 'string',
          description: 'Starting cell (e.g., "A1")',
          default: 'A1',
        },
      },
      required: ['file_path', 'sheet_name', 'data'],
    },
    handler: async (args: {
      file_path: string;
      sheet_name: string;
      data: any[][];
      start_cell?: string;
    }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      let worksheet = workbook.getWorksheet(args.sheet_name);
      if (!worksheet) {
        worksheet = workbook.addWorksheet(args.sheet_name);
      }
      
      const startCell = args.start_cell || 'A1';
      const [colLetter, rowStr] = startCell.match(/([A-Z]+)(\d+)/)?.slice(1) || ['A', '1'];
      const startRow = parseInt(rowStr);
      const startCol = colLetter.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
      
      args.data.forEach((row, rowIndex) => {
        row.forEach((value, colIndex) => {
          worksheet.getCell(startRow + rowIndex, startCol + colIndex).value = value;
        });
      });
      
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Data written to ${args.sheet_name} starting at ${startCell}`,
        rows_written: args.data.length,
      };
    },
  },
  
  {
    name: 'read_data_from_excel',
    description: 'Read data from an Excel worksheet',
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
          description: 'Cell range to read (e.g., "A1:D10")',
        },
      },
      required: ['file_path', 'sheet_name'],
    },
    handler: async (args: {
      file_path: string;
      sheet_name: string;
      range?: string;
    }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const worksheet = workbook.getWorksheet(args.sheet_name);
      if (!worksheet) {
        throw new Error(`Worksheet ${args.sheet_name} not found`);
      }
      
      let data: any[][] = [];
      
      if (args.range) {
        const [start, end] = args.range.split(':');
        const startMatch = start.match(/([A-Z]+)(\d+)/);
        const endMatch = end.match(/([A-Z]+)(\d+)/);
        
        if (startMatch && endMatch) {
          const startRow = parseInt(startMatch[2]);
          const endRow = parseInt(endMatch[2]);
          const startCol = startMatch[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
          const endCol = endMatch[1].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
          
          for (let row = startRow; row <= endRow; row++) {
            const rowData: any[] = [];
            for (let col = startCol; col <= endCol; col++) {
              const cell = worksheet.getCell(row, col);
              rowData.push(cell.value);
            }
            data.push(rowData);
          }
        }
      } else {
        worksheet.eachRow((row, rowNumber) => {
          const rowData: any[] = [];
          row.eachCell((cell, colNumber) => {
            rowData.push(cell.value);
          });
          data.push(rowData);
        });
      }
      
      return {
        success: true,
        data,
        rows: data.length,
        columns: data[0]?.length || 0,
      };
    },
  },
];