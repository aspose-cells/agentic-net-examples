// Title: C# – Refresh Parent Pivot Table and All Nested Child Pivot Tables with Aspose.Cells
// Description: Loads an Excel workbook, identifies the first worksheet's primary pivot table, refreshes its cache and calculations, retrieves every child pivot table via GetChildren(), updates each child’s data and layout, and saves the workbook. Includes file‑existence checks and robust error handling.
// Keywords: Aspose.Cells | C# | Refresh pivot tables | GetChildren method | nested pivot tables | shared cache | RefreshData | CalculateData | Excel automation | pivot table hierarchy
// Common Searches: Aspose.Cells refresh child pivot tables | C# get children pivot tables Aspose | how to update nested pivot tables in .NET | RefreshData CalculateData Aspose.Cells example | synchronize pivot table hierarchy programmatically
// Developer Intent: Programmatically update a parent pivot table and all its dependent child pivots so their data and layout stay in sync.
// Use Cases: After bulk data import, ensure the main pivot table and every linked child pivot reflect the new values before distribution. | Generate multi‑sheet reports where several pivot tables share a cache; refresh each child to prevent stale calculations. | Automate workbook validation in a CI pipeline, confirming no child pivot tables remain outdated after transformations.
// AI Prompts: Write C# code using Aspose.Cells to refresh a parent pivot table and all its child pivots in a workbook. | Explain the GetChildren() method for PivotTable objects in Aspose.Cells and how to handle an empty child collection. | Provide best‑practice error handling for refreshing nested pivot tables with Aspose.Cells, including file‑not‑found scenarios.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace RefreshNestedPivotTables
{
    // Loads an Excel workbook, identifies the first worksheet's primary pivot table, refreshes its cache and calculations, retrieves every child pivot table via GetChildren(), updates each child’s data and layout, and saves the workbook. Includes file‑existence checks and robust error handling.
    public class Program
    {
        public static void Main()
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                    return;
                }

                // Load the workbook that contains the parent pivot table
                Workbook workbook = new Workbook(inputFile);

                // Work with the first worksheet (adjust if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("Warning: No pivot tables found in the first worksheet.");
                    workbook.Save(outputFile);
                    Console.WriteLine($"Workbook saved as \"{outputFile}\".");
                    return;
                }

                // Assume the parent pivot table is the first one in the collection
                PivotTable parentPivot = worksheet.PivotTables[0];

                // Refresh the parent pivot table (optional but ensures base cache is up‑to‑date)
                parentPivot.RefreshData();
                parentPivot.CalculateData();

                // Retrieve all child pivot tables that use the parent pivot table as their data source
                PivotTable[] childPivots = parentPivot.GetChildren();

                // Refresh each child pivot table to synchronize its hierarchy with the parent
                if (childPivots != null)
                {
                    foreach (PivotTable child in childPivots)
                    {
                        child.RefreshData();      // Refresh data from the shared cache
                        child.CalculateData();    // Recalculate the pivot table layout
                    }
                }

                // Save the updated workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook refreshed and saved as \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
