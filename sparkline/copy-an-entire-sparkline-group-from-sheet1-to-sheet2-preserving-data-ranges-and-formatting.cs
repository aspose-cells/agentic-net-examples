// Title: Copy an entire sparkline group from Sheet1 to Sheet2 while preserving data ranges and formatting with Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy sparkline group from Sheet1 to Sheet2 preserving formatting | C# duplicate worksheet with sparklines using Aspose.Cells | How to keep sparkline data ranges when copying Excel sheets in .NET | Programmatically copy an Excel sheet including sparklines with Aspose.Cells | Copy worksheet and retain sparkline styles in C#
// Tags: worksheet.copy preserving sparklines Aspose.Cells | sparkline group duplication .NET | excel sheet clone with sparkline data ranges C# | retain sparkline formatting Aspose.Cells | copy worksheet including sparklines Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, ensures Sheet1 exists, creates or clears Sheet2, then uses Worksheet.Copy to duplicate all content—including every sparkline group with its data ranges and formatting—from Sheet1 to Sheet2, and finally saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get source (Sheet1) and destination (Sheet2) worksheets
            Worksheet srcSheet = workbook.Worksheets["Sheet1"];
            if (srcSheet == null)
            {
                Console.WriteLine("Error: Source worksheet \"Sheet1\" does not exist.");
                return;
            }

            Worksheet destSheet = workbook.Worksheets["Sheet2"];
            if (destSheet == null)
            {
                // Create Sheet2 if it does not exist
                destSheet = workbook.Worksheets.Add("Sheet2");
            }
            else
            {
                // Clear existing content to avoid mixing data
                destSheet.Cells.Clear();
                destSheet.SparklineGroups.Clear();
            }

            // Copy the entire content of Sheet1 (including sparklines) to Sheet2
            srcSheet.Copy(destSheet);

            // Save the workbook with the copied data
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
