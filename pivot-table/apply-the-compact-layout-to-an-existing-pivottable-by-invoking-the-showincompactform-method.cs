// Title: C# – Apply Compact Layout to a PivotTable with Aspose.Cells ShowInCompactForm
// Description: Loads an existing workbook, selects the first worksheet and its first PivotTable, switches the table to Compact layout using ShowInCompactForm, refreshes and recalculates the data, then saves the result as a new file.
// Keywords: Aspose.Cells ShowInCompactForm | compact layout pivot table C# | Aspose.Cells pivot table layout change | refresh pivot data Aspose.Cells | save workbook after pivot modification
// Common Searches: Aspose.Cells set pivot table to compact layout | ShowInCompactForm example C# | refresh pivot table after layout change Aspose.Cells | how to save workbook after modifying pivot table
// Developer Intent: Change a PivotTable to Compact layout and update its data programmatically using Aspose.Cells for .NET.
// Use Cases: Convert a standard PivotTable to Compact form before exporting reports. | Automate pivot layout adjustments in a batch workbook‑processing pipeline. | Ensure pivot calculations stay current after layout changes in generated spreadsheets.
// AI Prompts: Write C# code that opens a workbook, applies ShowInCompactForm to the first PivotTable, refreshes the data, and saves the file using Aspose.Cells. | Explain the impact of ShowInCompactForm on field arrangement and why RefreshData and CalculateData are required afterward. | Create robust error handling for scenarios where a worksheet contains no PivotTables when applying Compact layout.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCompactLayout
{
    // Loads an existing workbook, selects the first worksheet and its first PivotTable, switches the table to Compact layout using ShowInCompactForm, refreshes and recalculates the data, then saves the result as a new file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a pivot table
            Workbook workbook = new Workbook("input.xlsx");

            // Assume the pivot table is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the first pivot table in the worksheet
            if (sheet.PivotTables.Count > 0)
            {
                PivotTable pivotTable = sheet.PivotTables[0];

                // Apply the Compact layout to the pivot table
                pivotTable.ShowInCompactForm();

                // Refresh and recalculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();
            }
            else
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
