// Title: C# – Apply SlicerStyleLight1 to an Excel slicer with Aspose.Cells
// Description: Creates a workbook, builds a pivot table, adds a slicer linked to the pivot, sets the slicer’s StyleType to SlicerStyleLight1, and saves the file as an XLSX document.
// Keywords: Aspose.Cells slicer style C# | SlicerStyleLight1 example | Excel slicer formatting .NET | pivot table slicer Aspose | set slicer style programmatically
// Common Searches: how to set slicer style Aspose.Cells C# | SlicerStyleLight1 code sample | apply built‑in slicer style in .NET | change Excel slicer appearance with Aspose
// Developer Intent: Demonstrate how to assign the built‑in Light1 style to a slicer and persist the change when saving the workbook.
// Use Cases: Generate a dashboard workbook where slicers match a corporate Light theme. | Automate report creation with predefined slicer formatting for end‑user filtering. | Produce Excel files that include styled slicers to improve visual consistency across multiple sheets.
// AI Prompts: Show code to apply SlicerStyleDark1 to a slicer in Aspose.Cells C#. | How can I change the slicer style after loading an existing workbook? | List all SlicerStyleType values and give an example of selecting one based on a config setting.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerStyleDemo
{
    // Creates a workbook, builds a pivot table, adds a slicer linked to the pivot, sets the slicer’s StyleType to SlicerStyleLight1, and saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Fruit";
            cells["A2"].Value = "Apple";
            cells["A3"].Value = "Orange";
            cells["A4"].Value = "Banana";
            cells["B1"].Value = "Quantity";
            cells["B2"].Value = 10;
            cells["B3"].Value = 15;
            cells["B4"].Value = 20;

            // Add a pivot table based on the sample data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column
            pivot.AddFieldToArea(PivotFieldType.Data, 1); // Quantity column

            // Add a slicer linked to the pivot table's first field (Fruit)
            int slicerIdx = sheet.Slicers.Add(pivot, "G3", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Apply the built‑in light style 1 to the slicer
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Save the workbook with the styled slicer
            workbook.Save("SlicerStyleLight1Demo.xlsx");
        }
    }
}
