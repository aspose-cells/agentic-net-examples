// Title: Set PivotTable Layout to Tabular Form with Aspose.Cells for .NET (C#)
// Description: Loads an existing workbook, locates the first PivotTable, switches its layout to Tabular using ShowInTabularForm(), refreshes and recalculates the data, then saves the updated file.
// Keywords: Aspose.Cells C# PivotTable | Tabular layout PivotTable | ShowInTabularForm method | refresh pivot data | calculate pivot data | modify pivot programmatically | .NET Excel pivot table | change pivot layout
// Common Searches: Aspose.Cells change pivot table to tabular form C# | How to set Tabular layout for a PivotTable using Aspose.Cells | C# code to refresh and recalculate PivotTable after layout change | Find first PivotTable in workbook with Aspose.Cells | Save workbook after modifying PivotTable layout
// Developer Intent: Apply Tabular layout to an existing PivotTable in a .NET workbook.
// Use Cases: Prepare a report by converting a PivotTable to Tabular form before exporting to PDF for clearer column alignment. | Standardize data presentation across multiple worksheets by enforcing Tabular layout on each PivotTable. | Ensure calculated fields reflect the new structure by refreshing and recalculating the PivotTable after layout modification.
// AI Prompts: Generate C# code that iterates through all worksheets in a workbook, finds every PivotTable, applies ShowInTabularForm(), refreshes, recalculates, and saves the file. | Create a reusable method named SetFirstPivotToTabular(string inputPath, string outputPath) that returns a boolean indicating success. | Provide a C# example that detects the absence of PivotTables in a workbook and logs a friendly message without throwing an exception.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableTabularLayoutDemo
{
    // Loads an existing workbook, locates the first PivotTable, switches its layout to Tabular using ShowInTabularForm(), refreshes and recalculates the data, then saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through worksheets to find the first pivot table
            PivotTable pivotTable = null;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.PivotTables.Count > 0)
                {
                    // Get the first pivot table in this worksheet
                    pivotTable = sheet.PivotTables[0];
                    break;
                }
            }

            if (pivotTable == null)
            {
                Console.WriteLine("No pivot table found in the workbook.");
                return;
            }

            // Change the layout of the pivot table to Tabular form
            pivotTable.ShowInTabularForm();

            // Refresh and recalculate the pivot table to apply the layout change
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");

            Console.WriteLine("Pivot table layout changed to Tabular and workbook saved.");
        }
    }
}
