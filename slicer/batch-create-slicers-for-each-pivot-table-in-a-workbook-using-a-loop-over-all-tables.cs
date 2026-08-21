// Title: C# – Batch create slicers for all PivotTables in an Excel workbook with Aspose.Cells
// Description: Loads a workbook, loops through every worksheet and its PivotTables, adds a slicer for each table using the first base field, positions slicers with column offsets, sets a custom caption, and saves the file.
// Keywords: Aspose.Cells | C# | Excel slicer | PivotTable slicer | batch slicer creation | add slicer programmatically | multiple pivot tables | slicer placement | custom slicer caption | .NET Excel automation
// Common Searches: add slicer to every pivot table using Aspose.Cells C# | loop through worksheets and create slicers for pivot tables | programmatic slicer placement in Excel with Aspose | set custom caption for slicers in Aspose.Cells | batch generate slicers for Excel pivot tables .NET
// Developer Intent: Automatically generate a slicer for each PivotTable in a workbook.
// Use Cases: Add interactive filters to all PivotTables in a financial reporting workbook. | Standardize slicer layout across dashboard sheets that contain multiple PivotTables. | Export data with pre‑configured slicers and custom captions for end‑user analysis.
// AI Prompts: Generate C# code that uses Aspose.Cells to add slicers to every PivotTable in a workbook, arranging them side‑by‑side with unique captions. | Explain how to calculate dynamic slicer positions to avoid overlap when a worksheet has many PivotTables. | Show how to select a specific base field (instead of the first) for each slicer when creating them in a loop.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Loads a workbook, loops through every worksheet and its PivotTables, adds a slicer for each table using the first base field, positions slicers with column offsets, sets a custom caption, and saves the file.
class BatchSlicerCreator
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the collection of pivot tables on the current worksheet
            PivotTableCollection pivots = sheet.PivotTables;

            // Loop over each pivot table
            for (int p = 0; p < pivots.Count; p++)
            {
                PivotTable pivot = pivots[p];

                // Ensure the pivot table has at least one base field to create a slicer for
                if (pivot.BaseFields.Count == 0)
                    continue;

                // Determine a placement for the slicer (row and column indices)
                // Here we place slicers in the first rows, offsetting columns for each pivot table
                int slicerRow = 0;
                int slicerColumn = p * 5; // simple offset to avoid overlap

                // Add a slicer using the first base field of the pivot table
                // Using the overload: Add(PivotTable pivot, int row, int column, PivotField baseField)
                int slicerIndex = sheet.Slicers.Add(pivot, slicerRow, slicerColumn, pivot.BaseFields[0]);

                // Optional: customize the slicer (e.g., set caption)
                Slicer slicer = sheet.Slicers[slicerIndex];
                slicer.Caption = $"Slicer_{pivot.Name}";
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
