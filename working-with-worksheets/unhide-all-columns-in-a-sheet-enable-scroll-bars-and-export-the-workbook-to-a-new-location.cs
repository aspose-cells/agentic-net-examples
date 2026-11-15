// Title: Unhide every column in all worksheets, retain default scrollbars, and save the workbook to a new location with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells code to unhide all columns in each worksheet of an existing Excel file | How to save an Aspose.Cells workbook to a different folder after modifying column visibility | Ensuring output directory exists before calling Workbook.Save in Aspose.Cells C# | Preserving default scrollbars when exporting a workbook using Aspose.Cells for .NET
// Tags: unhide all columns Aspose.Cells C# | modify column visibility across worksheets Aspose.Cells | save workbook to new path Aspose.Cells | create output directory before workbook.Save C# | retain default scrollbars Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, iterates through every worksheet, sets each column's IsHidden property to false to unhide them, ensures the target output directory exists, and then saves the updated workbook to a new location while preserving the default scrollbars.
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
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Unhide all columns in the current sheet
                    int maxColumn = sheet.Cells.MaxColumn;
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        // Use the correct property to control visibility
                        sheet.Cells.Columns[col].IsHidden = false;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the modified workbook to a new location
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
