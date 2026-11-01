// Title: How to iterate merged cell areas and log each cell’s row and column indices with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that loads a workbook, retrieves all merged cell ranges from a worksheet, and prints the start and end row/column of each range together with the row and column of every cell inside the range. | Modify the example to output the merged area coordinates and individual cell indices to a CSV file instead of the console. | Extend the solution to loop through all worksheets in the workbook and generate a separate log for each sheet showing merged area details.
// Common Searches: Aspose.Cells C# enumerate merged cells and get individual cell positions | How to get row and column numbers of cells inside merged ranges using Aspose.Cells | C# example for GetMergedAreas and iterating over each cell in the range | Log merged area coordinates to a file with Aspose.Cells .NET | Iterate through all worksheets and list merged cell areas Aspose.Cells
// Tags: Aspose.Cells GetMergedAreas iteration | log merged cell coordinates C# | enumerate cells within merged range Aspose.Cells | process merged areas across worksheets .NET | output merged area details to CSV Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel workbook, uses Worksheet.Cells.GetMergedAreas() to obtain merged cell ranges, iterates each range to output its start/end row and column and the row/column of every cell inside, with optional file existence check and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve merged cell areas (returns CellArea[])
            CellArea[] mergedAreas = worksheet.Cells.GetMergedAreas();

            // Iterate through each merged area
            foreach (CellArea area in mergedAreas)
            {
                Console.WriteLine($"Mapped Area - StartRow: {area.StartRow}, StartColumn: {area.StartColumn}, EndRow: {area.EndRow}, EndColumn: {area.EndColumn}");

                // Optionally, log each individual cell within the area
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        Console.WriteLine($"Cell - Row: {row}, Column: {col}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
