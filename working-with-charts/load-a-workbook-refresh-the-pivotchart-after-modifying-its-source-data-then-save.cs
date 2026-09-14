// Title: Refresh a PivotChart after changing source data in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx file, modify a cell that feeds a PivotTable, refresh the PivotTable cache and its linked PivotChart, then save the workbook with Aspose.Cells in C#. | Generate C# code that verifies the input workbook, updates cell B2, calls PivotTable.RefreshData() and PivotTable.CalculateData(), and writes the result to a new Excel file. | Create a script that programmatically validates the source file, edits source data, forces a PivotTable refresh to update the associated chart, and outputs the updated workbook using Aspose.Cells.
// Common Searches: how to programmatically refresh a pivot chart after editing source data with Aspose.Cells C# | Aspose.Cells .NET refresh pivot table cache and linked chart after changing cell value | C# example to update Excel cell and recalculate pivot chart using Aspose.Cells library | refresh pivot chart in workbook without opening Excel using Aspose.Cells | Aspose.Cells RefreshData CalculateData methods for pivot tables
// Tags: Aspose.Cells pivot chart refresh C# | update source cell pivot table cache .NET | RefreshData method Aspose.Cells | CalculateData method Aspose.Cells | load workbook modify cell save Excel

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example checks for an input Excel file, loads it with Aspose.Cells, changes the value of cell B2 (the PivotTable source), refreshes the first PivotTable using RefreshData and CalculateData (which also updates any linked PivotChart), and saves the modified workbook to a new file while handling possible errors.
class PivotChartRefreshExample
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
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // ----- Modify source data -----
            // Example: change the value of cell B2 (source data for the pivot)
            worksheet.Cells["B2"].PutValue(12345);

            // ----- Refresh the PivotTable (which updates the linked PivotChart) -----
            // Assuming the first PivotTable on the sheet is the one used by the chart
            if (worksheet.PivotTables.Count > 0)
            {
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Refresh the data cache of the PivotTable
                pivotTable.RefreshData();

                // Recalculate the PivotTable values after the data refresh
                pivotTable.CalculateData();
            }
            else
            {
                Console.WriteLine("Warning: No PivotTable found on the first worksheet.");
            }

            // ----- Save the workbook with the refreshed PivotChart -----
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
