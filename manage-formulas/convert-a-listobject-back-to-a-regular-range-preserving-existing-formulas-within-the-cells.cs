// Title: Remove an Excel ListObject (table) while preserving cell formulas using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete Excel table but keep formulas | How to convert a ListObject to a normal range with Aspose.Cells | Remove first ListObject from worksheet while preserving formulas Aspose.Cells | C# Aspose.Cells keep cell formulas after removing a table | Convert Excel ListObject to range without losing formulas in .NET
// Tags: remove ListObject preserving formulas Aspose.Cells | convert Excel table to range Aspose.Cells | ListObjects.RemoveAt keep formulas | Aspose.Cells table deletion retain formulas | C# Excel ListObject to regular cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads a workbook, checks for a ListObject on the first worksheet, removes the first ListObject using ListObjects.RemoveAt (which leaves the underlying cell data and formulas intact), ensures the output directory exists, and saves the modified file, handling errors gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook containing the ListObject (table)
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one ListObject (table) on the sheet
            if (sheet.ListObjects.Count > 0)
            {
                // Remove the first ListObject while keeping its data.
                // Using RemoveAt on the collection avoids ambiguity with extension methods.
                sheet.ListObjects.RemoveAt(0);
            }
            else
            {
                Console.WriteLine("No ListObject (table) found on the worksheet.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
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
