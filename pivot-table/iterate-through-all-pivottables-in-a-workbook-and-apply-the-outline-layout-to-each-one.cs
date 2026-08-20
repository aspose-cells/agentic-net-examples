// Title: Set Outline Layout for All PivotTables in an Excel Workbook with Aspose.Cells (C#)
// Description: Loads a workbook, loops through every worksheet, retrieves each PivotTableCollection, applies the ShowInOutlineForm method to switch every pivot table to outline view, and saves the updated file.
// Keywords: Aspose.Cells C# pivot table outline | ShowInOutlineForm | apply outline layout to pivot tables | iterate pivot tables workbook | bulk pivot table formatting Aspose
// Common Searches: Aspose.Cells set outline layout for all pivot tables | C# iterate worksheets and apply ShowInOutlineForm | how to change pivot table view to outline using Aspose | bulk update pivot table layout in Excel with .NET
// Developer Intent: Apply the outline view to every pivot table in a workbook programmatically.
// Use Cases: Standardize reporting by converting all pivot tables to outline view before distribution. | Prepare workbooks for printing with a consistent outline layout for better readability. | Automate cleanup of multiple sheets to ensure uniform pivot table presentation.
// AI Prompts: Generate C# code that loads a workbook, iterates through all worksheets, and calls ShowInOutlineForm on each pivot table, handling sheets without pivots gracefully. | Provide an example that sets the outline layout for all pivot tables and also configures automatic subtotals and collapsed levels using Aspose.Cells APIs. | Create a reusable method that accepts a Workbook object, applies the outline form to every pivot table, and returns the modified workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotOutlineDemo
{
    // Loads a workbook, loops through every worksheet, retrieves each PivotTableCollection, applies the ShowInOutlineForm method to switch every pivot table to outline view, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of pivot tables on the current worksheet
                PivotTableCollection pivots = sheet.PivotTables;

                // Apply Outline layout to every pivot table found
                for (int i = 0; i < pivots.Count; i++)
                {
                    PivotTable pivot = pivots[i];
                    pivot.ShowInOutlineForm();   // Layout the pivot table in outline form
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
