// Title: Create a slicer linked to an existing pivot table using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, selects the first worksheet and its first pivot table, picks the first base field, adds a slicer at cell E2 that is bound to the pivot, customizes the caption and style, refreshes the slicer (and the pivot), and saves the file with the new slicer.
// Keywords: Aspose.Cells | C# | .NET | Excel slicer | add slicer to pivot table | linked slicer | Slicers.Add | pivot base field | slicer style | SlicerStyleLight2 | refresh slicer | save workbook with slicer | Aspose.Cells Slicers API
// Common Searches: Aspose.Cells add slicer to pivot table C# | How to link a slicer to a pivot table using Aspose.Cells | Set slicer caption and style with Aspose.Cells .NET | Refresh slicer after creation Aspose.Cells | Place slicer at specific cell location in Excel using Aspose.Cells
// Developer Intent: Insert a slicer that is bound to an existing pivot table and adjust its visual properties.
// Use Cases: Enable interactive filtering of pivot data by adding a slicer for the first base field. | Apply a consistent visual theme by setting the slicer caption and a light style. | Ensure the pivot reflects the latest data by refreshing the slicer before saving the workbook.
// AI Prompts: Generate C# code with Aspose.Cells that adds a slicer for a specified pivot field at cell B5 and applies the SlicerStyleDark1 style. | Show how to retrieve the slicer index after calling sheet.Slicers.Add and then modify its Caption and Style properties. | Provide robust error‑handling for cases where the worksheet has no pivot tables or the pivot has no base fields before creating a slicer.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Loads an Excel workbook, selects the first worksheet and its first pivot table, picks the first base field, adds a slicer at cell E2 that is bound to the pivot, customizes the caption and style, refreshes the slicer (and the pivot), and saves the file with the new slicer.
class SlicerLinkedToPivot
{
    static void Main()
    {
        // Load an existing workbook that already contains a pivot table
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the pivot table is on the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve the first pivot table in the worksheet
        if (sheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }
        PivotTable pivot = sheet.PivotTables[0];

        // Determine the field name to base the slicer on (use the first base field)
        if (pivot.BaseFields.Count == 0)
        {
            Console.WriteLine("Pivot table has no base fields to create a slicer.");
            return;
        }
        string baseFieldName = pivot.BaseFields[0].Name;

        // Add a slicer linked to the pivot table.
        // The slicer will be placed with its upper‑left corner at cell "E2".
        int slicerIndex = sheet.Slicers.Add(pivot, "E2", baseFieldName);

        // Access the newly created slicer to optionally set properties
        Slicer slicer = sheet.Slicers[slicerIndex];
        slicer.Caption = $"{baseFieldName} Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

        // Refresh the slicer (also refreshes the associated pivot table)
        slicer.Refresh();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
