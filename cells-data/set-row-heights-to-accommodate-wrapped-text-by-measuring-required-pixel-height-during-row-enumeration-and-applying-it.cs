// Title: How to auto‑fit Excel row heights for wrapped text using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that sets IsTextWrapped = true for every cell in a worksheet and then calls AutoFitRow for each row. | Update the logic to capture the height returned by GetRowHeight after AutoFitRow and apply SetRowHeight with that exact value. | Include a check that creates the destination folder if it does not already exist before saving the workbook.
// Common Searches: resize rows after text wrapping with Aspose.Cells | C# example to calculate required row height in points using Aspose.Cells | wrap text in all cells and adjust row sizes Aspose.Cells .NET | save workbook to new folder creating directory if missing Aspose.Cells C#
// Tags: auto‑fit row height Aspose.Cells C# | enable text wrap all cells Aspose.Cells | get row height points Aspose.Cells | create output directory before saving C# | programmatic Excel row height adjustment .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, enables text wrapping for every cell, auto‑fits each row, reads the calculated row height in points, explicitly sets that height, ensures the output folder exists, and saves the modified file.
class RowHeightAdjuster
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Iterate through rows that contain data
            for (int rowIndex = 0; rowIndex <= sheet.Cells.MaxDataRow; rowIndex++)
            {
                // Enable text wrapping for each cell in the current row
                for (int colIndex = 0; colIndex <= sheet.Cells.MaxDataColumn; colIndex++)
                {
                    var cell = sheet.Cells[rowIndex, colIndex];
                    var style = cell.GetStyle();
                    style.IsTextWrapped = true; // enable wrapping
                    cell.SetStyle(style);
                }

                // Auto‑fit the row height based on wrapped content
                sheet.AutoFitRow(rowIndex);

                // Retrieve the calculated height (in points) and re‑apply it
                double requiredPoints = sheet.Cells.GetRowHeight(rowIndex);
                sheet.Cells.SetRowHeight(rowIndex, requiredPoints);
            }

            // Ensure output directory exists
            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
