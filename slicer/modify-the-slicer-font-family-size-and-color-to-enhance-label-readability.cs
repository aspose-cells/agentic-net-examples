// Title: Aspose.Cells for .NET – Change Slicer Font Family, Size, and Color (C#)
// Description: C# example that creates a workbook, adds a pivot table, inserts a slicer, and customizes the slicer label by setting Font.Name, Font.Size, and Font.Color, with optional caption text, then saves the file.
// Keywords: Aspose.Cells slicer font | C# slicer label color | set slicer font size Aspose | change slicer font family .NET | pivot table slicer styling | Excel slicer customization | Aspose.Cells API example | GitHub Aspose.Cells slicer sample
// Common Searches: how to change slicer font in Aspose.Cells C# | set slicer label color and size .NET | customize slicer caption font Aspose.Cells | C# code to modify slicer font family Excel | Aspose.Cells example for slicer styling
// Developer Intent: Apply a specific font family, size, and color to a slicer’s label for better readability.
// Use Cases: Standardize corporate typography on slicer controls in automated Excel reports. | Enhance visual contrast of slicer items in dashboards for end‑user clarity. | Add a descriptive caption with styled text to guide slicer interaction.
// AI Prompts: Generate C# code using Aspose.Cells that sets a slicer’s font to Arial, size 14, red color, and hides the caption. | Show how to iterate over all slicers on a worksheet and apply the same Font.Name, Font.Size, and Font.Color. | Explain the steps to load an existing workbook, locate a slicer by name, and modify its Font properties with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerFontCustomization
{
    // C# example that creates a workbook, adds a pivot table, inserts a slicer, and customizes the slicer label by setting Font.Name, Font.Size, and Font.Color, with optional caption text, then saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["B4"].Value = 150;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Amount field
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table for the "Category" field
            int slicerIdx = sheet.Slicers.Add(pivot, "F1", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Customize the slicer label font: family, size, and color
            // The slicer’s visual representation is a Shape; its Font property can be modified directly
            slicer.Shape.Font.Name = "Calibri";          // Font family
            slicer.Shape.Font.Size = 12;                // Font size (points)
            slicer.Shape.Font.Color = Color.DarkBlue;   // Font color

            // Optionally, make the caption visible and set its text
            slicer.ShowCaption = true;
            slicer.Caption = "Select Category";

            // Save the workbook with the customized slicer
            workbook.Save("SlicerFontCustomization.xlsx");
        }
    }
}
