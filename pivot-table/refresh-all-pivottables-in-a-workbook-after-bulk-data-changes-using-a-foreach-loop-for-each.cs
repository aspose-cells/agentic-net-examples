// Title: Refresh every PivotTable in an Excel workbook after bulk data changes using Aspose.Cells for .NET (C#)
// Common Searches: C# code to refresh all pivot tables in an Excel file after modifying data with Aspose.Cells | how to loop through worksheets and update pivot tables programmatically using Aspose.Cells .NET | Aspose.Cells RefreshData and CalculateData for multiple pivot tables example | bulk data update Excel workbook then refresh pivot tables with Aspose.Cells C# | save workbook after refreshing pivot tables using Aspose.Cells for .NET
// Tags: Aspose.Cells pivot table data refresh C# | loop over worksheets to refresh pivot tables Aspose | RefreshData and CalculateData API Aspose.Cells | apply bulk data changes Excel workbook Aspose | export updated workbook after pivot refresh .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, optionally performs bulk data modifications, iterates through each worksheet and its PivotTables, calls RefreshData and CalculateData on every PivotTable to update and recalculate the data, and finally saves the refreshed workbook to a new file.
class RefreshPivotTables
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

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // TODO: Perform bulk data changes here
            // Example: workbook.Worksheets[0].Cells["A2"].PutValue(123);

            // Refresh all PivotTables in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    // Refresh the data source of the PivotTable
                    pivotTable.RefreshData();

                    // Recalculate the PivotTable after refreshing data
                    pivotTable.CalculateData();
                }
            }

            // Save the updated workbook to a new file
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
