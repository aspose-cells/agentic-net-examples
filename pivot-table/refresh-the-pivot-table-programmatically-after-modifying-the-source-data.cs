// Title: Programmatically refresh an existing PivotTable after modifying source data with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# refresh pivot table after updating worksheet cells | How to recalculate a PivotTable programmatically using Aspose.Cells .NET | RefreshData and CalculateData methods for PivotTable in Aspose.Cells C# example | Update source data and refresh pivot table in an Excel file with Aspose.Cells | C# code to refresh pivot table when source data changes using Aspose.Cells
// Tags: Aspose.Cells RefreshData API | Aspose.Cells CalculateData API | programmatic pivot table refresh .NET | modify worksheet cell Aspose.Cells | pivot table data source update C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads a workbook (or creates one), changes a cell value on the first worksheet, checks for an existing PivotTable, and calls RefreshData() followed by CalculateData() to synchronize the PivotTable with the updated source before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Load the workbook if the file exists; otherwise create a new workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Optional: add some default data to avoid empty workbook issues.
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Value");
                ws.Cells["A2"].PutValue(0);
            }

            // Access the first worksheet (assumed to contain source data).
            Worksheet dataSheet = workbook.Worksheets[0];

            // Modify source data as needed (example: change value in cell A2).
            dataSheet.Cells["A2"].PutValue(12345);

            // Ensure a pivot table exists before attempting to refresh it.
            if (dataSheet.PivotTables.Count > 0)
            {
                PivotTable pivot = dataSheet.PivotTables[0];
                // Refresh the pivot table to reflect the updated source data.
                pivot.RefreshData();      // Refreshes the data source.
                pivot.CalculateData();    // Recalculates the pivot table.
            }
            else
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
            }

            // Save the updated workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
