// Title: Change the data source range of an existing pivot table in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file using Aspose.Cells, locates the first pivot table on the first worksheet, and updates its data source to a specified range. | Show how to call PivotTable.ChangeDataSource to replace the source range of a pivot cache and then save the modified workbook to a new file. | Provide a robust C# example that verifies the input file exists, changes the pivot table source range, creates missing output directories, and handles any exceptions.
// Common Searches: aspnet aspose.cells how to change pivot table source range in C# | c# update pivot cache data source for existing Excel workbook using Aspose.Cells | load workbook modify first pivot table source range aspose.cells | change pivot table data source to new range and save workbook with Aspose.Cells .NET | example code for Aspose.Cells ChangeDataSource method C#
// Tags: aspose.cells pivot table source range change | c# aspose.cells update pivot cache | excel workbook modify pivot source | aspose.cells pivot table source update example | c# load workbook edit pivot table

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable classes

// The sample checks for an input.xlsx file, loads it into an Aspose.Cells Workbook, accesses the first worksheet, retrieves its first PivotTable, changes the pivot's data source to "Sheet1!A1:D100" using ChangeDataSource, ensures the output directory exists, saves the workbook as output.xlsx, and includes exception handling for robustness.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is on the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of pivot tables on this worksheet
            PivotTableCollection pivotTables = worksheet.PivotTables;

            // Ensure there is at least one pivot table to work with
            if (pivotTables.Count > 0)
            {
                // Retrieve the first pivot table (or locate by name if known)
                PivotTable pivot = pivotTables[0];

                // Define the new data source range.
                // Format: "SheetName!StartCell:EndCell"
                // Example changes the source to cells A1 through D100 on Sheet1.
                string[] newDataSource = new string[] { "Sheet1!A1:D100" };

                try
                {
                    // Apply the new data source to the pivot table.
                    // This updates the pivot cache to reference the new range.
                    pivot.ChangeDataSource(newDataSource);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to change pivot data source: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No pivot tables found in the workbook.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook to a new file (or overwrite the original)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
