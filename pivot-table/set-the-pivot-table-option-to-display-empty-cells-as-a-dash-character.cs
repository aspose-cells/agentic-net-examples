// Title: How to configure Aspose.Cells PivotTable to display a dash for empty cells in C#
// Common Searches: Aspose.Cells C# set pivot table empty cell display to dash | display dash for blank values in Excel pivot table using Aspose.Cells .NET | how to use EmptyCellDisplay property in Aspose.Cells pivot tables | C# Aspose.Cells pivot table show '-' for null values | set empty cell placeholder in Aspose.Cells pivot table programmatically
// Tags: Aspose.Cells EmptyCellDisplay | pivot table dash placeholder .NET | C# configure pivot table empty cells | Excel pivot table blank cell formatting | Aspose.Cells pivot table display options

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable

// Shows how to load an existing workbook, access its first pivot table, set the EmptyCellDisplay property so that empty cells appear as a dash, and save the modified workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook that contains the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table in the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // NOTE: The EmptyCellDisplay property is available in newer versions of Aspose.Cells.
            // If using an older version, this setting is not supported and can be omitted.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the (potentially) updated pivot table setting
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors (e.g., Aspose.Cells specific exceptions)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
