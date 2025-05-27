export interface Tool {
  name: string;
  description: string;
  inputSchema: {
    type: 'object';
    properties: Record<string, any>;
    required?: string[];
  };
  handler: (args: any) => Promise<any>;
}

export interface WorkbookMetadata {
  path: string;
  sheets: string[];
  activeSheet?: string;
  properties?: Record<string, any>;
}

export interface CellFormat {
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
  };
  fill?: {
    type?: 'pattern' | 'gradient';
    pattern?: string;
    color?: string;
    bgColor?: string;
  };
  alignment?: {
    horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify';
    vertical?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
  };
  border?: {
    top?: BorderStyle;
    bottom?: BorderStyle;
    left?: BorderStyle;
    right?: BorderStyle;
  };
  numFmt?: string;
}

export interface BorderStyle {
  style?: 'thin' | 'medium' | 'thick' | 'double';
  color?: string;
}

export interface ChartOptions {
  type: 'column' | 'bar' | 'line' | 'pie' | 'area' | 'scatter';
  title?: string;
  xAxis?: {
    title?: string;
  };
  yAxis?: {
    title?: string;
  };
  legend?: {
    position?: 'top' | 'bottom' | 'left' | 'right';
  };
}