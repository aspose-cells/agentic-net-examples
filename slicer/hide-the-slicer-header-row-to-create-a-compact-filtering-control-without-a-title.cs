// Title: Hide Slicer Header (Caption) in Aspose.Cells for .NET to Create a Compact Filter Control
// Description: Shows how to create a workbook, add sample data, build a pivot table, insert a slicer linked to the "Category" field, and remove the slicer header by setting Slicer.ShowCaption = false. The sample also applies a light style and saves the file as SlicerHeaderHidden.xlsx.
// Keywords: Aspose.Cells | .NET | C# slicer | hide slicer header | remove slicer caption | ShowCaption false | compact slicer | Excel pivot table slicer | slicer styling | Excel automation
// Common Searches: Aspose.Cells hide slicer header | remove slicer caption C# | compact slicer without title Aspose | ShowCaption property Excel slicer | how to hide slicer title in generated Excel
// Developer Intent: Remove the slicer caption to produce a space‑saving slicer control.
// Use Cases: Design dashboards where screen space is limited and slicer titles are unnecessary. | Generate automated reports with multiple slicers that need a uniform, minimal appearance. | Customize slicer styling after hiding the caption to match a clean, minimalist layout.
// AI Prompts: Write C# code using Aspose.Cells to add a slicer to a pivot table and hide its caption for a compact UI. | Show how to set Slicer.ShowCaption = false and apply a style to the slicer in Aspose.Cells for .NET. | Provide an example that creates a workbook, builds a pivot table, inserts a slicer, hides the header row, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerHeaderHideDemo
{
    // Shows how to create a workbook, add sample data, build a pivot table, insert a slicer linked to the "Category" field, and remove the slicer header by setting Slicer.ShowCaption = false. The sample also applies a light style and saves the file as SlicerHeaderHidden.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "Fruit";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Vegetable";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Grain";
            cells["B4"].Value = 50;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D5", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.CalculateData();

            // Add a slicer linked to the pivot table for the "Category" field
            int slicerIdx = sheet.Slicers.Add(pivot, "F5", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Hide the slicer header (caption) to make it compact
            slicer.ShowCaption = false;

            // Optional: adjust slicer appearance
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.NumberOfColumns = 1;

            // Save the workbook
            workbook.Save("SlicerHeaderHidden.xlsx");
        }
    }
}
