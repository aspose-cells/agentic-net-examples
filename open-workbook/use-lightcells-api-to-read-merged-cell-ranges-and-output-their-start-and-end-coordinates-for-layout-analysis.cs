// Title: Read merged cell ranges with Aspose.Cells (C#) – LightCells API example
// Description: Loads an Excel workbook, extracts every merged area using GetMergedAreas(), converts zero‑based indices to A1 notation with CellsHelper, and prints each region's start and end coordinates for layout analysis.
// Keywords: Aspose.Cells merged cells C# | LightCells read merged areas | GetMergedAreas .NET | Excel merged range coordinates | layout analysis Aspose.Cells
// Common Searches: list merged cell ranges Aspose.Cells C# | convert merged area indices to A1 notation | retrieve merged cell coordinates with LightCells | how to enumerate merged cells in .NET Excel library
// Developer Intent: Obtain all merged cell ranges from a worksheet and output their start/end addresses.
// Use Cases: Generate PDF layout preserving merged regions; build a mapping of merged cells for custom reporting; validate merged areas before data import; perform visual analysis of spreadsheet structure.
// AI Prompts: Show how to modify the sample to stream merged cells with LightCells for very large workbooks. | Create a method that returns a List<string> of merged area addresses instead of writing to the console. | Explain handling of merged cells when exporting a worksheet to HTML using Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel workbook, extracts every merged area using GetMergedAreas(), converts zero‑based indices to A1 notation with CellsHelper, and prints each region's start and end coordinates for layout analysis.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be analyzed
        string inputFile = "input.xlsx";

        // Load the workbook (standard loading; LightCells can be used for large files,
        // but merged area information is available after loading the worksheet).
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve all merged cell areas in the worksheet
        CellArea[] mergedAreas = worksheet.Cells.GetMergedAreas();

        // Output the total number of merged areas
        Console.WriteLine($"Number of merged areas: {mergedAreas.Length}");

        // Iterate through each merged area and display its start and end coordinates
        foreach (CellArea area in mergedAreas)
        {
            // Convert zero‑based indices to the usual Excel A1 style for readability
            string startAddress = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
            string endAddress   = CellsHelper.CellIndexToName(area.EndRow,   area.EndColumn);

            Console.WriteLine($"Merged area: {startAddress} ({area.StartRow},{area.StartColumn}) " +
                              $"to {endAddress} ({area.EndRow},{area.EndColumn})");
        }
    }
}
