// Title: C# – Stream Large Excel Worksheet Row‑by‑Row with Aspose.Cells LightCellsDataHandler
// Description: Learn how to read a massive Excel file in C# using Aspose.Cells LightCells. By attaching a custom LightCellsDataHandler to LoadOptions, rows and cells are processed sequentially, eliminating the need to load the whole workbook into memory.
// Keywords: Aspose.Cells | LightCells | LightCellsDataHandler | C# | stream large Excel | row streaming | memory‑efficient Excel processing | read Excel without loading | large worksheet | cell iteration
// Common Searches: Aspose.Cells stream rows C# | LightCellsDataHandler example .NET | read large Excel file without loading into memory | process Excel rows on the fly Aspose.Cells | C# memory‑efficient Excel reading
// Developer Intent: Read and handle rows and cells of a huge worksheet sequentially while keeping memory usage low.
// Use Cases: Log or analyze every cell value in a multi‑gigabyte Excel file without OOM errors. | Filter, transform, or aggregate rows during streaming before exporting to another format. | Copy selected rows to a new workbook or CSV while the source file remains unmaterialized.
// AI Prompts: Create a LightCellsDataHandler that writes each processed row to a CSV file during streaming. | Show how to stop LightCells processing after row 10,000 using the handler methods. | Provide code to correctly handle merged cells and formulas while streaming rows with LightCells.

using System;
using Aspose.Cells;

// Learn how to read a massive Excel file in C# using Aspose.Cells LightCells. By attaching a custom LightCellsDataHandler to LoadOptions, rows and cells are processed sequentially, eliminating the need to load the whole workbook into memory.
class Program
{
    static void Main()
    {
        // Path to the large Excel file to be streamed
        string inputPath = "LargeData.xlsx";

        // Create an instance of the custom LightCellsDataHandler
        var handler = new StreamingHandler();

        // Configure LoadOptions to use the handler
        var loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = handler;

        // Load the workbook in LightCells (streaming) mode.
        // The workbook is not fully materialized in memory.
        var workbook = new Workbook(inputPath, loadOptions);

        // At this point all rows and cells have been processed by the handler.
        // If you need to save a copy, you can do so (optional):
        // workbook.Save("Copy.xlsx");
    }

    // Custom implementation of LightCellsDataHandler that streams rows and cells.
    class StreamingHandler : LightCellsDataHandler
    {
        // Called before processing a worksheet.
        public bool StartSheet(Worksheet sheet)
        {
            Console.WriteLine($"Processing sheet: {sheet.Name}");
            return true; // Continue processing this sheet.
        }

        // Called before processing each row.
        public bool StartRow(int rowIndex)
        {
            // Return true to process the row and its cells.
            return true;
        }

        // Called after the row object is created; can be used to read row properties.
        public bool ProcessRow(Row row)
        {
            Console.WriteLine($"Row {row.Index}:");
            // Return true to also process cells in this row.
            return true;
        }

        // Called before processing each cell in the current row.
        public bool StartCell(int columnIndex)
        {
            // Return true to process the cell.
            return true;
        }

        // Called for each cell that should be processed.
        public bool ProcessCell(Cell cell)
        {
            // Output cell address and its value.
            Console.WriteLine($"  {cell.Name} = {cell.Value}");
            return true;
        }
    }
}
