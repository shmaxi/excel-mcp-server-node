# Excel MCP Server (Node.js)

Excel MCP Server for manipulating Excel files - Node.js implementation with npx support.

## Installation

You can run this server using npx without installing:

```bash
npx @shmaxi/excel-mcp-server stdio
```

Or install globally:

```bash
npm install -g @shmaxi/excel-mcp-server
excel-mcp-server stdio
```

## Usage

### Stdio Transport (Local Integration)

```json
{
  "mcpServers": {
    "excel": {
      "command": "npx",
      "args": ["@shmaxi/excel-mcp-server", "stdio"]
    }
  }
}
```

### SSE Transport (Remote Connections)

```bash
FASTMCP_PORT=8000 npx @shmaxi/excel-mcp-server sse
```

## Available Tools

### Workbook Operations
- `create_workbook` - Create a new Excel workbook
- `get_workbook_metadata` - Get metadata about an Excel workbook

### Data Operations
- `write_data_to_excel` - Write data to a worksheet
- `read_data_from_excel` - Read data from a worksheet

### Formatting Operations
- `format_range` - Apply formatting to a range of cells
- `merge_cells` - Merge a range of cells
- `unmerge_cells` - Unmerge previously merged cells
- `apply_formula` - Apply a formula to a cell

### Chart Operations
- `create_chart` - Create a chart in the worksheet
- `create_pivot_table` - Create a pivot table

### Worksheet Operations
- `create_worksheet` - Create a new worksheet
- `copy_worksheet` - Copy a worksheet
- `delete_worksheet` - Delete a worksheet
- `rename_worksheet` - Rename a worksheet

### Range Operations
- `copy_range` - Copy a range of cells
- `delete_range` - Delete a range of cells
- `validate_excel_range` - Validate if a range reference is valid
- `validate_formula_syntax` - Validate Excel formula syntax

## Development

```bash
# Install dependencies
npm install

# Build
npm run build

# Run in development mode
npm run dev stdio
```

## License

MIT