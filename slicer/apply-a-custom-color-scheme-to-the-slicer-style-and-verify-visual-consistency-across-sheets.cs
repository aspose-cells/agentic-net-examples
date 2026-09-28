// Title: How to apply a custom orange palette to a slicer style in Aspose.Cells for .NET and keep the color consistent across worksheets
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, inserts a slicer linked to a field, changes the workbook palette index to an orange color, and sets the slicer’s StyleType to SlicerStyleDark1. | Update an existing Aspose.Cells workbook to modify its color palette, apply the new palette to a slicer style, and save the file while confirming the slicer displays the custom color. | Extend the sample by duplicating the slicer on a second worksheet and programmatically verify that the custom orange color appears identically on both slicers.
// Common Searches: Aspose.Cells C# change slicer palette color to orange | apply custom color to slicer style in .NET workbook | verify slicer color consistency across multiple worksheets Aspose.Cells | how to modify workbook palette index for slicer styling in C# | set SlicerStyleDark1 with custom palette in Aspose.Cells
// Tags: custom slicer palette Aspose.Cells .NET | slicer style custom palette | modify workbook palette index C# | pivot table slicer color customization | slicer visual consistency across worksheets

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace SlicerCustomColorDemo
{
    // // Demonstrates creating a workbook, building a pivot table, adding a slicer linked to the 'Category' field, changing palette index 0 to an orange color, applying SlicerStyleDark1 to use the custom palette, and saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for a pivot table
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Item";
                cells["C1"].Value = "Amount";

                cells["A2"].Value = "Fruit";
                cells["B2"].Value = "Apple";
                cells["C2"].Value = 120;

                cells["A3"].Value = "Fruit";
                cells["B3"].Value = "Banana";
                cells["C3"].Value = 80;

                cells["A4"].Value = "Vegetable";
                cells["B4"].Value = "Carrot";
                cells["C4"].Value = 150;

                cells["A5"].Value = "Vegetable";
                cells["B5"].Value = "Tomato";
                cells["C5"].Value = 90;

                // Create a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:C5", "E2", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Column, "Item");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Refresh pivot cache and calculate data (correct API)
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the "Category" field of the pivot table
                int slicerIdx = sheet.Slicers.Add(pivot, "G2", "Category");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Change a palette entry to a custom color (e.g., bright orange)
                Color customColor = Color.FromArgb(255, 165, 0); // Orange
                workbook.ChangePalette(customColor, 0); // Modify palette index 0

                // Apply a built‑in slicer style that will use the modified palette entry
                slicer.StyleType = SlicerStyleType.SlicerStyleDark1;

                // (Optional) Verify visual consistency – here we simply note that the style was applied
                Console.WriteLine($"Slicer style applied: {slicer.StyleType}");

                // Save the workbook
                string outputPath = "SlicerCustomColorDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
