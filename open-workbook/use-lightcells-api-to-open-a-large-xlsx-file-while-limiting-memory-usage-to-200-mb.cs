// Title: Open a large XLSX workbook with Aspose.Cells LightCells API while capping memory at 200 MB (C#)
// Description: Demonstrates how to load a multi‑gigabyte XLSX file in LightCells mode using a SimpleLightCellsHandler, configure LoadOptions with MemorySetting.FileCache, and keep RAM usage near 200 MB. The example reads the first ten cells of column A and optionally saves the processed workbook.
// Keywords: Aspose.Cells | LightCells API | MemorySetting.FileCache | large XLSX low memory | C# .NET | row‑by‑row processing | file‑based cache | memory‑constrained Excel | Open large workbook | 200 MB limit
// Common Searches: Aspose.Cells LightCells open huge XLSX with memory limit | C# load large Excel file using file cache | How to restrict Aspose.Cells memory usage to 200 MB | LightCellsDataHandler example for low‑memory scenarios | Process multi‑GB workbook without high RAM consumption
// Developer Intent: The developer needs to read or manipulate a very large XLSX file in .NET while ensuring the application does not exceed roughly 200 MB of RAM.
// Use Cases: Extract specific rows or columns from a multi‑gigabyte spreadsheet on a server with limited RAM. | Convert or export a massive workbook to another format in a memory‑restricted environment. | Run data validation or transformation on huge Excel files in batch jobs without OOM errors.
// AI Prompts: Write C# code that opens a 4 GB XLSX file with Aspose.Cells LightCells, uses MemorySetting.FileCache, and iterates through every row. | Create a custom LightCellsDataHandler that skips empty rows and logs processed cell counts to reduce memory overhead. | Explain how to set a custom temporary directory for the file‑cache used by LightCells when processing large workbooks.

using System;
using Aspose.Cells;

// Demonstrates how to load a multi‑gigabyte XLSX file in LightCells mode using a SimpleLightCellsHandler, configure LoadOptions with MemorySetting.FileCache, and keep RAM usage near 200 MB. The example reads the first ten cells of column A and optionally saves the processed workbook.
class Program
{
    static void Main()
    {
        // Create a LightCellsDataHandler that simply allows processing of all sheets, rows, and cells.
        var handler = new SimpleLightCellsHandler();

        // Configure load options to use file cache mode, which keeps memory usage low (≈200 MB or less).
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.MemorySetting = MemorySetting.FileCache;      // Use file‑based cache.
        loadOptions.LightCellsDataHandler = handler;              // Attach the handler.

        // Load the large XLSX file in LightCells mode with the specified options.
        Workbook workbook = new Workbook("LargeFile.xlsx", loadOptions);

        // Example: read and display the first 10 values from column A of the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];
        for (int row = 0; row < 10; row++)
        {
            Console.WriteLine(sheet.Cells[row, 0].StringValue);
        }

        // Save the workbook after processing (optional).
        workbook.Save("ProcessedLargeFile.xlsx");
    }
}

// Implementation of LightCellsDataHandler that permits full processing without custom logic.
public class SimpleLightCellsHandler : LightCellsDataHandler
{
    // Called when a worksheet is about to be processed.
    public bool StartSheet(Worksheet sheet)
    {
        // Return true to continue processing this sheet.
        return true;
    }

    // Called before a row is processed.
    public bool StartRow(int rowIndex)
    {
        // Return true to process this row.
        return true;
    }

    // Called after a row is prepared; return true to process its cells.
    public bool ProcessRow(Row row)
    {
        return true;
    }

    // Called before a cell is processed.
    public bool StartCell(int columnIndex)
    {
        // Return true to process this cell.
        return true;
    }

    // Called for each cell; no custom handling needed.
    public bool ProcessCell(Cell cell)
    {
        // Simply continue processing.
        return true;
    }
}
