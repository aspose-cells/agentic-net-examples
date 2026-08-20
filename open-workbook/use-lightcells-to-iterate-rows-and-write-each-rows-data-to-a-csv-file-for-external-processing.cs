// Title: C# Export Excel to CSV with Aspose.Cells LightCellsDataHandler (row‑by‑row streaming)
// Description: Demonstrates a custom CsvExportHandler that inherits from LightCellsDataHandler to stream each worksheet row directly to a CSV file, handling commas, quotes and line breaks while keeping memory usage low.
// Keywords: Aspose.Cells LightCellsDataHandler | C# CSV export | stream Excel to CSV | row by row processing | large worksheet export | .NET Excel to CSV | memory‑efficient CSV generation
// Common Searches: How to export Excel to CSV using LightCells in Aspose.Cells C# | LightCellsDataHandler example for CSV output | Stream rows from a workbook to a CSV file with Aspose.Cells | C# write Excel cells to CSV without loading whole file
// Developer Intent: Generate a CSV file from an Excel workbook by iterating rows with LightCellsDataHandler, avoiding full workbook loading.
// Use Cases: Export massive worksheets to CSV without exhausting memory. | Create a real‑time CSV pipeline for downstream analytics or ETL processes. | Apply custom escaping rules for commas, quotes and new‑line characters during export.
// AI Prompts: Write a LightCellsDataHandler in C# that streams worksheet rows to a CSV file with proper escaping. | Show how to configure LoadOptions.LightCellsDataHandler to export an Excel file to CSV without fully loading it. | Explain how to modify the CsvExportHandler to skip empty rows and write a single header line.

using System;
using System.IO;
using Aspose.Cells;

// Custom LightCellsDataHandler that writes each processed cell to a CSV file.
// Demonstrates a custom CsvExportHandler that inherits from LightCellsDataHandler to stream each worksheet row directly to a CSV file, handling commas, quotes and line breaks while keeping memory usage low.
class CsvExportHandler : LightCellsDataHandler
{
    private readonly StreamWriter _writer;
    private bool _firstRow = true;
    private int _currentColumn = -1;

    public CsvExportHandler(string outputPath)
    {
        _writer = new StreamWriter(outputPath);
    }

    // Process all sheets.
    public bool StartSheet(Worksheet sheet) => true;

    // Called before processing a row.
    public bool StartRow(int rowIndex)
    {
        // Write line break before every row except the first.
        if (!_firstRow)
            _writer.WriteLine();
        else
            _firstRow = false;

        _currentColumn = -1;
        return true; // Continue processing this row.
    }

    // Not used for CSV export, just continue.
    public bool ProcessRow(Row row) => true;

    // Called before each cell in the current row.
    public bool StartCell(int columnIndex)
    {
        _currentColumn = columnIndex;
        return true; // Process this cell.
    }

    // Write cell value to CSV, handling commas and quotes.
    public bool ProcessCell(Cell cell)
    {
        if (_currentColumn > 0)
            _writer.Write(",");

        string value = cell.StringValue ?? string.Empty;

        // Escape double quotes.
        if (value.Contains("\""))
            value = value.Replace("\"", "\"\"");

        // Enclose in quotes if needed.
        if (value.Contains(",") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r"))
            value = $"\"{value}\"";

        _writer.Write(value);
        return true; // Keep processing.
    }

    // Flush and close the writer when done.
    public void Close()
    {
        _writer.Flush();
        _writer.Dispose();
    }
}

class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample workbook with some data.
        // -------------------------------------------------
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Name");
        ws.Cells["B1"].PutValue("Age");
        ws.Cells["A2"].PutValue("John");
        ws.Cells["B2"].PutValue(30);
        ws.Cells["A3"].PutValue("Alice");
        ws.Cells["B3"].PutValue(25);

        // Save the workbook to a temporary file (required for loading with LightCells).
        string tempPath = "temp.xlsx";
        wb.Save(tempPath, SaveFormat.Xlsx);

        // -------------------------------------------------
        // 2. Export the workbook to CSV using LightCellsDataHandler.
        // -------------------------------------------------
        string csvPath = "output.csv";
        var handler = new CsvExportHandler(csvPath);

        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = handler;

        // Loading triggers the handler; data is written to CSV during this call.
        Workbook loadedWb = new Workbook(tempPath, loadOptions);

        // Finalize CSV file.
        handler.Close();

        Console.WriteLine($"CSV file has been created at: {csvPath}");
    }
}
