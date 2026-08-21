// Title: Set Slicer Layout to Right‑to‑Left (RTL) in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the "Category" field, sets the slicer’s TextDirection to RightToLeft, optionally switches the worksheet to right‑to‑left display, and saves the file as SlicerRTL.xlsx.
// Keywords: Aspose.Cells | C# | slicer RTL | right to left slicer | TextDirectionType.RightToLeft | pivot table slicer | display worksheet right to left | Arabic Excel report | Hebrew Excel slicer | Aspose.Cells API
// Common Searches: Aspose.Cells set slicer right to left | C# Aspose.Cells RTL slicer | how to make slicer layout RTL in .NET | right‑to‑left worksheet with slicer Aspose | pivot table slicer Arabic layout
// Developer Intent: Configure a slicer’s text direction to right‑to‑left for RTL language support in an Aspose.Cells workbook.
// Use Cases: Generate localized Excel reports for Arabic or Hebrew users with a pivot‑table slicer that reads from right to left. | Apply RTL layout to both the slicer and the worksheet to match right‑to‑left UI conventions. | Programmatically set the slicer’s TextDirection before saving the workbook to ensure correct rendering on all devices.
// AI Prompts: Write C# code using Aspose.Cells to add a slicer to a pivot table and set its TextDirection to RightToLeft. | Explain how to enable right‑to‑left display for a worksheet and its slicers in Aspose.Cells. | Provide a step‑by‑step guide for creating an RTL slicer for a pivot table in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;

namespace AsposeCellsSlicerRTL
{
    // Creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the "Category" field, sets the slicer’s TextDirection to RightToLeft, optionally switches the worksheet to right‑to‑left display, and saves the file as SlicerRTL.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Product";
                cells["A2"].Value = "Fruits";
                cells["B2"].Value = "Apple";
                cells["A3"].Value = "Fruits";
                cells["B3"].Value = "Banana";
                cells["A4"].Value = "Vegetables";
                cells["B4"].Value = "Carrot";

                // Create a pivot table based on the data
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Product");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table (field "Category")
                // Note: The Add method expects destination cell first, then field name
                int slicerIdx = sheet.Slicers.Add(pivot, "E6", "Category");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Set the slicer layout direction to right‑to‑left
                slicer.Shape.TextDirection = TextDirectionType.RightToLeft;

                // Optionally, set the whole worksheet to RTL display as well
                sheet.DisplayRightToLeft = true;

                // Save the workbook
                workbook.Save("SlicerRTL.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
