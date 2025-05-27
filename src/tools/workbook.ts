import ExcelJS from 'exceljs';
import { Tool, WorkbookMetadata } from '../types/index.js';
import path from 'path';
import fs from 'fs/promises';

export const workbookTools: Tool[] = [
  {
    name: 'create_workbook',
    description: 'Create a new Excel workbook',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path where the workbook will be saved',
        },
        sheet_name: {
          type: 'string',
          description: 'Name of the initial worksheet',
          default: 'Sheet1',
        },
      },
      required: ['file_path'],
    },
    handler: async (args: { file_path: string; sheet_name?: string }) => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet(args.sheet_name || 'Sheet1');
      
      // Ensure directory exists
      const dir = path.dirname(args.file_path);
      await fs.mkdir(dir, { recursive: true });
      
      await workbook.xlsx.writeFile(args.file_path);
      
      return {
        success: true,
        message: `Workbook created at ${args.file_path}`,
        path: args.file_path,
      };
    },
  },
  
  {
    name: 'get_workbook_metadata',
    description: 'Get metadata about an Excel workbook',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: {
          type: 'string',
          description: 'Path to the Excel file',
        },
      },
      required: ['file_path'],
    },
    handler: async (args: { file_path: string }): Promise<WorkbookMetadata> => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const sheets = workbook.worksheets.map(ws => ws.name);
      
      return {
        path: args.file_path,
        sheets,
        activeSheet: workbook.worksheets[0]?.name,
        properties: {
          creator: workbook.creator,
          lastModifiedBy: workbook.lastModifiedBy,
          created: workbook.created,
          modified: workbook.modified,
        },
      };
    },
  },
];