// Title: Read merged cell ranges with Aspose.Cells LightCells API in C# – start and end coordinates
// Description: C# example that uses Aspose.Cells LightCells API to open an Excel workbook, enumerate all merged cell areas via Cells.MergedCells, convert the zero‑based row/column indices to standard Excel addresses with CellsHelper, and print each range's start and end coordinates for layout analysis.
// Keywords: Aspose.Cells LightCells merged cells | C# read merged cell ranges | Excel merged cell coordinates | CellsHelper address conversion | list merged areas Aspose
// Common Searches: Aspose.Cells LightCells get merged cell ranges C# | how to list merged cells in Excel using Aspose | convert merged cell indices to A1 address Aspose.Cells | enumerate merged areas in a worksheet with Aspose | C# code to output start and end of merged cells
// Developer Intent: Retrieve every merged cell range from a worksheet and display its start and end cell addresses.
// Use Cases: Map layout sections before exporting a spreadsheet to PDF or image. | Validate template structure by confirming required merged regions exist. | Generate a CSV or JSON map of merged areas for custom rendering pipelines.
// AI Prompts: Create a C# method that returns a list of CellArea objects representing all merged ranges in a given worksheet using Aspose.Cells LightCells. | Write code to export each merged range's start and end A1 addresses to a CSV file with Aspose.Cells. | Show how to handle merged cells when converting a worksheet to PDF with Aspose.Cells, preserving the original layout.

using System;
using System.IO;
using Aspose.Cells;

// C# example that uses Aspose.Cells LightCells API to open an Excel workbook, enumerate all merged cell areas via Cells.MergedCells, convert the zero‑based row/column indices to standard Excel addresses with CellsHelper, and print each range's start and end coordinates for layout analysis.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be analyzed
        string filePath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook – suitable for retrieving merged cells
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve all merged cell areas from the worksheet via Cells.MergedCells
            var mergedAreas = worksheet.Cells.MergedCells;

            Console.WriteLine($"Number of merged areas: {mergedAreas.Count}");
            foreach (CellArea area in mergedAreas)
            {
                // Convert zero‑based row/column indices to Excel cell addresses for readability
                string startAddress = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                string endAddress   = CellsHelper.CellIndexToName(area.EndRow,   area.EndColumn);

                Console.WriteLine(
                    $"Merged area: {startAddress}:{endAddress} " +
                    $"(StartRow={area.StartRow}, StartColumn={area.StartColumn}, " +
                    $"EndRow={area.EndRow}, EndColumn={area.EndColumn})");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
