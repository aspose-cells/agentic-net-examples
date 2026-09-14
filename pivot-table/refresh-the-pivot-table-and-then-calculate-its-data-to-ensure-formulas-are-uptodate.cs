// Title: Refresh a PivotTable and recalculate its data in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, locate the first PivotTable, invoke RefreshData() and CalculateData(), then save the file using Aspose.Cells in C#. | Write C# code that verifies worksheets and PivotTables exist before calling RefreshData and CalculateData on a PivotTable with Aspose.Cells. | Show how to handle missing input files and runtime exceptions while refreshing and recalculating a PivotTable via Aspose.Cells.
// Common Searches: Aspose.Cells C# how to refresh pivot table source data and recalculate formulas | example code to call RefreshData and CalculateData on a PivotTable using Aspose.Cells .NET | update pivot table after changing source data programmatically with Aspose.Cells in C#
// Tags: pivot table data refresh Aspose.Cells C# | pivot table calculation Aspose.Cells C# | load and save Excel workbook Aspose.Cells | check worksheet and pivot table existence Aspose.Cells | exception handling Aspose.Cells file operations

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshAndCalculate
{
    // The sample loads input.xlsx, ensures a worksheet and at least one PivotTable are present, refreshes the first PivotTable's source data with RefreshData(), recalculates its formulas using CalculateData(), and saves the updated workbook to output.xlsx while handling missing files and other exceptions.
    public class Program
    {
        public static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook that contains a pivot table
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook has at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any worksheets.");
                    return;
                }

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the first worksheet.");
                    return;
                }

                // Get the first pivot table in the worksheet
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Refresh the pivot table's data from its source
                pivotTable.RefreshData();

                // Calculate the pivot table data so that formulas are up‑to‑date
                pivotTable.CalculateData();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
