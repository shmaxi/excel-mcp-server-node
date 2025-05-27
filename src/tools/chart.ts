import ExcelJS from 'exceljs';
import { Tool, ChartOptions } from '../types/index.js';

export const chartTools: Tool[] = [
  {
    name: 'create_chart',
    description: 'Create a chart in the worksheet',
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
        chart_type: {
          type: 'string',
          description: 'Type of chart (column, bar, line, pie, area, scatter)',
          enum: ['column', 'bar', 'line', 'pie', 'area', 'scatter'],
        },
        data_range: {
          type: 'string',
          description: 'Data range for the chart (e.g., "A1:B10")',
        },
        position: {
          type: 'object',
          description: 'Chart position',
          properties: {
            cell: {
              type: 'string',
              description: 'Top-left cell for the chart',
            },
            width: {
              type: 'number',
              description: 'Chart width in pixels',
              default: 600,
            },
            height: {
              type: 'number',
              description: 'Chart height in pixels',
              default: 400,
            },
          },
          required: ['cell'],
        },
        options: {
          type: 'object',
          description: 'Additional chart options',
          properties: {
            title: { type: 'string' },
            xAxis: {
              type: 'object',
              properties: {
                title: { type: 'string' },
              },
            },
            yAxis: {
              type: 'object',
              properties: {
                title: { type: 'string' },
              },
            },
            legend: {
              type: 'object',
              properties: {
                position: {
                  type: 'string',
                  enum: ['top', 'bottom', 'left', 'right'],
                },
              },
            },
          },
        },
      },
      required: ['file_path', 'sheet_name', 'chart_type', 'data_range', 'position'],
    },
    handler: async (args: {
      file_path: string;
      sheet_name: string;
      chart_type: string;
      data_range: string;
      position: {
        cell: string;
        width?: number;
        height?: number;
      };
      options?: ChartOptions;
    }) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(args.file_path);
      
      const worksheet = workbook.getWorksheet(args.sheet_name);
      if (!worksheet) {
        throw new Error(`Worksheet ${args.sheet_name} not found`);
      }
      
      // Note: ExcelJS has limited chart support
      // This is a placeholder implementation
      // In a real implementation, you might need to use a different library
      // or create charts through XML manipulation
      
      return {
        success: true,
        message: `Chart creation requested. Note: ExcelJS has limited chart support. Consider using alternative libraries for advanced charting.`,
        chart_type: args.chart_type,
        data_range: args.data_range,
        position: args.position,
      };
    },
  },
  
  {
    name: 'create_pivot_table',
    description: 'Create a pivot table',
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
          description: 'Source data range',
        },
        target_sheet: {
          type: 'string',
          description: 'Target worksheet for pivot table',
        },
        target_cell: {
          type: 'string',
          description: 'Target cell for pivot table',
          default: 'A1',
        },
        rows: {
          type: 'array',
          description: 'Fields for rows',
          items: { type: 'string' },
        },
        columns: {
          type: 'array',
          description: 'Fields for columns',
          items: { type: 'string' },
        },
        values: {
          type: 'array',
          description: 'Fields for values',
          items: {
            type: 'object',
            properties: {
              field: { type: 'string' },
              function: {
                type: 'string',
                enum: ['sum', 'count', 'average', 'min', 'max'],
              },
            },
          },
        },
      },
      required: ['file_path', 'source_sheet', 'source_range', 'target_sheet'],
    },
    handler: async (args: any) => {
      // Note: ExcelJS doesn't have built-in pivot table support
      // This is a placeholder implementation
      
      return {
        success: true,
        message: `Pivot table creation requested. Note: ExcelJS doesn't have built-in pivot table support. Consider using alternative libraries or Excel automation.`,
        source: `${args.source_sheet}!${args.source_range}`,
        target: `${args.target_sheet}!${args.target_cell || 'A1'}`,
      };
    },
  },
];