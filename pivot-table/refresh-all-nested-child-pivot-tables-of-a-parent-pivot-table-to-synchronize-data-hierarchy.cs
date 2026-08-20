// Title: Refresh All Nested Child Pivot Tables with Aspose.Cells for .NET (C#)
// Description: C# example that loads a workbook, validates worksheets and pivot tables, refreshes the parent pivot, retrieves its child pivots via GetChildren(), refreshes and recalculates each child, and saves the updated file. Includes basic error handling for missing files and empty sheets.
// Keywords: Aspose.Cells refresh child pivot tables | GetChildren pivot table .NET | nested pivot table refresh C# | RefreshData CalculateData Aspose | pivot table hierarchy synchronization | C# Excel pivot cache update | Aspose.Cells example refresh multiple pivots
// Common Searches: how to refresh child pivot tables in Aspose.Cells | Aspose.Cells GetChildren method example | C# refresh all pivot tables sharing a cache | update nested pivot tables programmatically | Aspose.Cells refresh data and recalculate pivots
// Developer Intent: Programmatically refresh a parent pivot table and all its linked child pivots to keep the hierarchy consistent.
// Use Cases: After modifying source data, ensure every dependent pivot reflects the changes before distribution. | Automate nightly Excel processing where several pivots share a cache and must stay synchronized. | Generate a refreshed report by loading an existing workbook, updating the parent and child pivots, and saving a new version.
// AI Prompts: Write C# code using Aspose.Cells to refresh a parent pivot table and all its child pivots, handling missing files and empty worksheets. | Explain the GetChildren method in Aspose.Cells and show how to iterate over the returned PivotTable array to recalculate each child. | Suggest enhancements for error handling and performance when refreshing a large number of nested pivot tables in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace RefreshNestedPivotTables
{
    // C# example that loads a workbook, validates worksheets and pivot tables, refreshes the parent pivot, retrieves its child pivots via GetChildren(), refreshes and recalculates each child, and saves the updated file. Includes basic error handling for missing files and empty sheets.
    public class Program
    {
        public static void Main()
        {
            try
            {
                const string inputFile = "InputWorkbook.xlsx";
                const string outputFile = "RefreshedWorkbook.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                    return;
                }

                // Load the workbook that contains the parent pivot table and its child pivot tables
                Workbook workbook = new Workbook(inputFile);

                // Ensure the workbook has at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("Error: The workbook does not contain any worksheets.");
                    return;
                }

                // Assume the parent pivot table is in the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("Error: No pivot tables found in the first worksheet.");
                    return;
                }

                // Get the first pivot table as the parent pivot
                PivotTable parentPivot = worksheet.PivotTables[0];

                // Refresh the parent pivot table (optional but ensures base cache is up‑to‑date)
                parentPivot.RefreshData();
                parentPivot.CalculateData();

                // Retrieve all child pivot tables that use the parent pivot table as their data source
                PivotTable[] childPivots = parentPivot.GetChildren();

                // Refresh each child pivot table to synchronize its hierarchy with the parent
                if (childPivots != null)
                {
                    foreach (PivotTable childPivot in childPivots)
                    {
                        childPivot.RefreshData();      // Refresh data from the shared cache
                        childPivot.CalculateData();    // Recalculate the displayed data
                    }
                }

                // Save the updated workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook refreshed and saved as \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
