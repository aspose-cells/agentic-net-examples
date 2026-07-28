// Title: C# – Export Workbook to CSV with LightCells API (no full model load) – Aspose.Cells
// Description: Shows how to stream cell values via a custom LightCellsDataProvider, save a workbook as a temporary XLSX using OoxmlSaveOptions, and convert it to CSV with ConversionUtility, enabling low‑memory export in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | LightCells API | C# CSV export | Low memory Excel | OoxmlSaveOptions | LightCellsDataProvider | ConversionUtility | stream workbook to CSV | export large workbook | without loading model
// Common Searches: Aspose.Cells LightCells export CSV C# | convert Excel to CSV without loading workbook | low memory CSV export Aspose.Cells | LightCellsDataProvider example | save workbook with LightCells then to CSV
// Developer Intent: Create a CSV file from an Excel workbook while keeping memory usage minimal by leveraging the LightCells API.
// Use Cases: Processing massive Excel files in data pipelines where memory is limited. | Generating CSV reports on the fly without persisting the full workbook in memory. | Customizing row and column selection during export via a user‑defined LightCellsDataProvider. | Integrating Aspose.Cells into serverless functions that require a low‑memory footprint.
// AI Prompts: Write a C# snippet that uses LightCellsDataProvider to export only columns A and C to CSV. | Explain how to add a row filter to SimpleLightCellsDataProvider before CSV conversion. | Show how to convert an in‑memory workbook stream directly to CSV with Aspose.Cells, avoiding temporary files. | Provide guidance on handling large worksheets with LightCells to prevent OutOfMemory exceptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Shows how to stream cell values via a custom LightCellsDataProvider, save a workbook as a temporary XLSX using OoxmlSaveOptions, and convert it to CSV with ConversionUtility, enabling low‑memory export in Aspose.Cells for .NET.
class ExportToCsvViaLightCells
{
    static void Main()
    {
        // Create a workbook and populate sample data
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Name");
        ws.Cells["B1"].PutValue("Age");
        ws.Cells["A2"].PutValue("John");
        ws.Cells["B2"].PutValue(30);
        ws.Cells["A3"].PutValue("Alice");
        ws.Cells["B3"].PutValue(25);

        // Save the workbook using LightCellsDataProvider to avoid loading the full model into memory
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
        {
            LightCellsDataProvider = new SimpleLightCellsDataProvider(wb)
        };
        string tempXlsx = "temp.xlsx";
        wb.Save(tempXlsx, saveOptions);

        // Convert the saved XLSX file to CSV without loading the model in this code
        string csvPath = "output.csv";
        ConversionUtility.Convert(tempXlsx, csvPath);

        Console.WriteLine("Exported to CSV: " + csvPath);
    }
}

// Simple LightCellsDataProvider that streams data from an existing workbook
class SimpleLightCellsDataProvider : LightCellsDataProvider
{
    private readonly Workbook _workbook;
    private int _currentRow = -1;
    private int _currentCell = -1;

    public SimpleLightCellsDataProvider(Workbook workbook)
    {
        _workbook = workbook;
    }

    // Start processing the first sheet (only one sheet in this example)
    public bool StartSheet(int sheetIndex)
    {
        return sheetIndex == 0;
    }

    // Return the next row index to be saved, or -1 when no more rows
    public int NextRow()
    {
        _currentRow++;
        _currentCell = -1;
        return _currentRow <= _workbook.Worksheets[0].Cells.MaxDataRow ? _currentRow : -1;
    }

    // No special row initialization needed
    public void StartRow(Row row) { }

    // Return the next cell (column) index to be saved, or -1 when no more cells in the row
    public int NextCell()
    {
        _currentCell++;
        return _currentCell <= _workbook.Worksheets[0].Cells.MaxDataColumn ? _currentCell : -1;
    }

    // Fill the cell with the value from the source workbook
    public void StartCell(Cell cell)
    {
        Cell srcCell = _workbook.Worksheets[0].Cells[_currentRow, _currentCell];
        if (srcCell != null && srcCell.Type != CellValueType.IsNull)
        {
            cell.PutValue(srcCell.Value);
        }
    }

    // Do not gather strings into a global pool
    public bool IsGatherString()
    {
        return false;
    }
}
