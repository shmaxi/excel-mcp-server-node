import ExcelJS from 'exceljs';
import { Tool } from '../types/index.js';

export const worksheetTools: Tool[] = [
  {
    name: 'create_worksheet',
    description: 'Create a new worksheet in an existing workbook',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheet_name: {
          type: 'string',
          description: 'Name for the new worksheet',
        },
      },
      required: ['file_path', 'sheet_name'],
    },
    handler: async (args: { file_path: string; sheet_name: string }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const existingSheet = workbook.getWorksheet(args.sheet_name);
      if (existingSheet) {
        throw new Error(`Worksheet ${args.sheet_name} already exists`);
      }
      
      workbook.addWorksheet(args.sheet_name);
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Worksheet ${args.sheet_name} created`,
      };
    },
  },
  
  {
    name: 'copy_worksheet',
    description: 'Copy a worksheet within the same workbook',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        source_sheet: {
          type: 'string',
          description: 'Name of the worksheet to copy',
        },
        target_sheet: {
          type: 'string',
          description: 'Name for the copied worksheet',
        },
      },
      required: ['file_path', 'source_sheet', 'target_sheet'],
    },
    handler: async (args: {
      file_path: string;
      source_sheet: string;
      target_sheet: string;
    }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const sourceWorksheet = workbook.getWorksheet(args.source_sheet);
      if (!sourceWorksheet) {
        throw new Error(`Source worksheet ${args.source_sheet} not found`);
      }
      
      const targetWorksheet = workbook.addWorksheet(args.target_sheet);
      
      // Copy all cells
      sourceWorksheet.eachRow((row, rowNumber) => {
        const newRow = targetWorksheet.getRow(rowNumber);
        row.eachCell((cell, colNumber) => {
          const newCell = newRow.getCell(colNumber);
          newCell.value = cell.value;
          newCell.style = cell.style;
        });
      });
      
      // Copy column properties
      sourceWorksheet.columns.forEach((col, index) => {
        if (col.width) {
          targetWorksheet.getColumn(index + 1).width = col.width;
        }
      });
      
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Worksheet ${args.source_sheet} copied to ${args.target_sheet}`,
      };
    },
  },
  
  {
    name: 'delete_worksheet',
    description: 'Delete a worksheet from a workbook',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        sheet_name: {
          type: 'string',
          description: 'Name of the worksheet to delete',
        },
      },
      required: ['file_path', 'sheet_name'],
    },
    handler: async (args: { file_path: string; sheet_name: string }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const worksheetId = workbook.getWorksheet(args.sheet_name)?.id;
      if (!worksheetId) {
        throw new Error(`Worksheet ${args.sheet_name} not found`);
      }
      
      workbook.removeWorksheet(worksheetId);
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Worksheet ${args.sheet_name} deleted`,
      };
    },
  },
  
  {
    name: 'rename_worksheet',
    description: 'Rename a worksheet',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Excel file',
        },
        old_name: {
          type: 'string',
          description: 'Current name of the worksheet',
        },
        new_name: {
          type: 'string',
          description: 'New name for the worksheet',
        },
      },
      required: ['file_path', 'old_name', 'new_name'],
    },
    handler: async (args: {
      file_path: string;
      old_name: string;
      new_name: string;
    }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const worksheet = workbook.getWorksheet(args.old_name);
      if (!worksheet) {
        throw new Error(`Worksheet ${args.old_name} not found`);
      }
      
      worksheet.name = args.new_name;
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Worksheet renamed from ${args.old_name} to ${args.new_name}`,
      };
    },
  },
];