// Title: Refresh Parent Pivot Table and Recalculate Dependent Child Pivots with Aspose.Cells for .NET
// Description: Loads an Excel workbook, refreshes the first (parent) pivot table, uses GetDependentPivotTables to locate all child pivots, refreshes and recalculates each child, then saves the updated file. Demonstrates correct total propagation after source data changes.
// Keywords: Aspose.Cells refresh pivot | parent pivot table refresh .NET | GetDependentPivotTables example | recalculate child pivots | Excel pivot table automation | C# Aspose.Cells pivot update | refresh all pivots in workbook
// Common Searches: how to refresh a parent pivot and its child pivots using Aspose.Cells | Aspose.Cells GetDependentPivotTables C# | refresh and calculate all pivot tables in an Excel file | update dependent pivot tables after data change .NET | Aspose.Cells refresh data for multiple pivots
// Developer Intent: Programmatically refresh a parent pivot table and then refresh and recalculate every child pivot that depends on it.
// Use Cases: After modifying source data via code, ensure the master pivot and all drill‑down pivots display accurate totals before exporting. | Automate nightly reporting where a primary pivot drives several linked child pivots; the routine keeps the entire report consistent. | Integrate into a data‑processing pipeline that validates pivot calculations by refreshing the parent and each dependent child pivot prior to saving.
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, refreshes the first pivot table, retrieves its dependent child pivots using GetDependentPivotTables, refreshes each child, and saves the file. | Explain the role of GetDependentPivotTables in Aspose.Cells and how to handle cases where no dependent pivots are returned. | Provide performance tips for refreshing and recalculating many pivot tables in a large workbook with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, refreshes the first (parent) pivot table, uses GetDependentPivotTables to locate all child pivots, refreshes and recalculates each child, then saves the updated file. Demonstrates correct total propagation after source data changes.
    public class RefreshParentAndChildrenPivotTables
    {
        public static void Run()
        {
            const string inputFile = "ParentPivotWorkbook.xlsx";
            const string outputFile = "ParentPivotWorkbook_Updated.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                    return;
                }

                // Load the workbook containing the parent pivot table and its children
                Workbook workbook = new Workbook(inputFile);

                // Assume the parent pivot table is the first one on the first worksheet
                Worksheet parentSheet = workbook.Worksheets[0];
                if (parentSheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                    return;
                }

                PivotTable parentPivot = parentSheet.PivotTables[0];

                // Refresh the parent pivot table data and recalculate
                parentPivot.RefreshData();
                parentPivot.CalculateData();

                // Retrieve all child pivot tables that depend on the parent pivot table
                PivotTable[] childPivots = parentPivot.GetDependentPivotTables();

                // Refresh each child pivot table's data and recalculate
                foreach (PivotTable childPivot in childPivots)
                {
                    childPivot.RefreshData();
                    childPivot.CalculateData();
                }

                // Save the updated workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully as '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
