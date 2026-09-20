// Title: Unmerge the merged range B2:C2 in an Excel workbook, copy its original value to D2, and save the file with Aspose.Cells for .NET
// AI Prompts: Locate the merged range that includes B2, unmerge it, read the original cell value, write that value into D2, and save the workbook. | Using Aspose.Cells, programmatically break a merged cell, duplicate its content to a neighboring cell, and export the updated spreadsheet.
// Common Searches: Aspose.Cells C# unmerge a specific merged range and copy its value to another cell | how to retrieve value from a merged cell after unmerging with Aspose.Cells .NET | save a modified Excel file to a new location after changing merged cells using Aspose.Cells | C# code to unmerge B2:C2 and duplicate its content to D2 in an existing workbook
// Tags: unmerge merged range Aspose.Cells .NET | duplicate cell value after unmerge Aspose.Cells | save updated workbook Aspose.Cells | handle merged cells programmatically C#

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing Excel file, finds and unmerges the merged range B2:C2, copies the original value to cell D2, ensures the output directory exists, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the merged range B2:C2 (zero‑based indices: row 1, columns 1‑2)
            int startRow = 1;      // B2 row index
            int startColumn = 1;   // B2 column index

            // Find the merged range that contains the specified cell
            Aspose.Cells.Range mergedRange = null;
            foreach (CellArea area in sheet.Cells.MergedCells)
            {
                if (area.StartRow <= startRow && startRow <= area.EndRow &&
                    area.StartColumn <= startColumn && startColumn <= area.EndColumn)
                {
                    int totalRows = area.EndRow - area.StartRow + 1;
                    int totalColumns = area.EndColumn - area.StartColumn + 1;
                    mergedRange = sheet.Cells.CreateRange(area.StartRow, area.StartColumn, totalRows, totalColumns);
                    break;
                }
            }

            // Unmerge the range if it exists
            mergedRange?.UnMerge();

            // Retrieve the original value from the former merged cell (now B2)
            object originalValue = sheet.Cells["B2"].Value;

            // Copy the value to the adjacent cell on the right (D2)
            sheet.Cells["D2"].Value = originalValue;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for debugging
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
