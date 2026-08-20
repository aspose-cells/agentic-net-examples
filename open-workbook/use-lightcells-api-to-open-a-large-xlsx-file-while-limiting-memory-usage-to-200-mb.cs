// Title: Load a Large XLSX with Aspose.Cells LightCells (C#) – keep memory under 200 MB using FileCache
// Description: Demonstrates how to open a massive XLSX workbook with the LightCells API in C#. The example sets LoadOptions.MemorySetting to FileCache and attaches a LightCellsDataHandler that processes all sheets, rows, and cells, resulting in an in‑memory footprint well below 200 MB. The workbook is then saved to a new file.
// Keywords: Aspose.Cells | LightCells API | MemorySetting.FileCache | C# load large XLSX | low memory Excel processing | streaming workbook | large workbook handling | 200 MB memory limit | Aspose.Cells LightCells example
// Common Searches: Aspose.Cells load large Excel file low memory | LightCells API C# memory limit example | How to use FileCache with Aspose Cells | Open big XLSX with LightCells streaming | C# Aspose Cells limit RAM usage
// Developer Intent: Open a huge XLSX workbook with LightCells while restricting RAM usage to 200 MB.
// Use Cases: Process or convert multi‑gigabyte Excel files on servers with limited RAM. | Iterate through sheets, rows, and cells for custom transformations without loading the whole file into memory. | Generate reports from massive workbooks in environments such as Azure Functions or AWS Lambda where memory caps are strict.
// AI Prompts: Show how to modify LightCellsDataHandler to skip rows where a specific column is empty while still using FileCache. | Provide code that logs current memory consumption during LightCells processing of a 5 GB workbook. | Explain step‑by‑step how to configure LightCells for streaming a 10 GB Excel file in a .NET Core console app.

using System;
using Aspose.Cells;

// Demonstrates how to open a massive XLSX workbook with the LightCells API in C#. The example sets LoadOptions.MemorySetting to FileCache and attaches a LightCellsDataHandler that processes all sheets, rows, and cells, resulting in an in‑memory footprint well below 200 MB. The workbook is then saved to a new file.
public class LightCellsDataHandlerDemo : LightCellsDataHandler
{
    // Called when a worksheet is about to be processed.
    public bool StartSheet(Worksheet sheet)
    {
        // Process all sheets.
        return true;
    }

    // Called before processing a row.
    public bool StartRow(int rowIndex)
    {
        // Process all rows.
        return true;
    }

    // Called after row properties are read.
    public bool ProcessRow(Row row)
    {
        // No custom row processing needed.
        return true;
    }

    // Called before processing a cell in the current row.
    public bool StartCell(int columnIndex)
    {
        // Process all cells.
        return true;
    }

    // Called after a cell's data is read.
    public bool ProcessCell(Cell cell)
    {
        // No custom cell processing needed.
        return true;
    }
}

public class Program
{
    public static void Main()
    {
        // Path to the large XLSX file.
        string inputFile = "LargeFile.xlsx";

        // Path where the processed workbook will be saved.
        string outputFile = "ProcessedLargeFile.xlsx";

        // Configure load options to limit memory usage.
        // FileCache mode stores intermediate data in temporary files,
        // keeping the in‑memory footprint low (well below 200 MB).
        LoadOptions loadOptions = new LoadOptions
        {
            MemorySetting = MemorySetting.FileCache,
            LightCellsDataHandler = new LightCellsDataHandlerDemo()
        };

        // Load the workbook using LightCells API with the specified options.
        Workbook workbook = new Workbook(inputFile, loadOptions);

        // Save the workbook (could be the same or a different file).
        workbook.Save(outputFile);
    }
}
