// Title: How to change the source data range of an existing PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change pivot table source range programmatically | Update PivotTable data source to a new range with Aspose.Cells .NET | Refresh a PivotTable after modifying its source data using Aspose.Cells | Example of ChangeDataSource method for an existing PivotTable in C# | Modify the source data of a PivotTable in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells ChangeDataSource pivot table | update pivot table source range .NET | refresh pivot table after data source change C# | modify existing pivot table source Aspose.Cells | pivot table source data update Excel workbook C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing workbook, checks for a PivotTable on the first worksheet, changes its source data range to "Sheet1!A1:D200" using the ChangeDataSource method, refreshes the PivotTable to apply the new data, and saves the updated workbook to a new file.
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found in the worksheet.");
                return;
            }

            // Retrieve the first PivotTable (change index if another is required)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Define the new source data range (e.g., Sheet1!A1:D200)
            string newSourceData = "Sheet1!A1:D200";

            // Update the source data range of the PivotTable
            // ChangeDataSource expects a string array in this API version
            pivotTable.ChangeDataSource(new[] { newSourceData });

            // Refresh the PivotTable so it reflects the new source data
            pivotTable.RefreshData();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated PivotTable
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
