import { Tool } from '../types/index.js';
import { workbookTools } from './workbook.js';
import { dataTools } from './data.js';
import { formatTools } from './format.js';
import { chartTools } from './chart.js';
import { worksheetTools } from './worksheet.js';
import { rangeTools } from './range.js';

export function registerTools(): Map<string, Tool> {
  const tools = new Map<string, Tool>();
  
  // Register all tool categories
  const allTools = [
    ...workbookTools,
    ...dataTools,
    ...formatTools,
    ...chartTools,
    ...worksheetTools,
    ...rangeTools,
  ];

  for (const tool of allTools) {
    tools.set(tool.name, tool);
  }

  return tools;
}