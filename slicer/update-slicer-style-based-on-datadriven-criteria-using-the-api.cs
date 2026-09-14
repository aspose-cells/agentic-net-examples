// Title: Apply a conditional built-in style to an Excel slicer based on total sales using Aspose.Cells for .NET
// AI Prompts: Create a workbook, add a pivot table from a data range, insert a slicer linked to the Product field, compute the sum of the Sales column, and set the slicer’s StyleType to SlicerStyleDark2 or SlicerStyleLight1 depending on whether the total exceeds a threshold. | Update the slicer’s caption, number of columns, width, and height after the conditional style has been applied. | Replace the built-in style logic with a custom style file or a different sales threshold, and adjust the code to select the appropriate style dynamically.
// Common Searches: how to set slicer style programmatically with Aspose.Cells in C# | conditional formatting of Excel slicer based on pivot table totals using Aspose | Aspose.Cells C# apply dark or light slicer style after calculating sales sum | change slicer appearance dynamically in .NET workbook with Aspose.Cells | example of using SlicerStyleType enumeration in Aspose.Cells
// Tags: conditional slicer styling Aspose.Cells | set slicer StyleType C# | pivot table total calculation Aspose.Cells | dynamic slicer appearance .NET | built-in slicer style enumeration | apply slicer style based on data threshold

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerStyleDynamicDemo
{
    // Demonstrates creating a workbook, building a pivot table, adding a slicer linked to the Product field, calculating total sales from the source range, and applying either a dark or light built-in slicer style based on a sales threshold before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for a pivot table
            // Columns: Product, Region, Sales
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Apple";
            cells["B2"].Value = "North";
            cells["C2"].Value = 120;

            cells["A3"].Value = "Apple";
            cells["B3"].Value = "South";
            cells["C3"].Value = 80;

            cells["A4"].Value = "Banana";
            cells["B4"].Value = "North";
            cells["C4"].Value = 150;

            cells["A5"].Value = "Banana";
            cells["B5"].Value = "South";
            cells["C5"].Value = 70;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Configure the pivot: Product as row, Region as column, Sales as data (sum)
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Column, "Region");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the "Product" field of the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "G3", "Product");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Determine a data‑driven condition:
            // If total sales > 300, use a dark style; otherwise, use a light style.
            double totalSales = 0;
            // The pivot data starts at the cell where the pivot is placed (E3)
            // The aggregated values are in the data area; we can sum them directly.
            // For simplicity, iterate over the source data range.
            for (int row = 2; row <= 5; row++) // rows 2‑5 contain sales values
            {
                totalSales += cells[row - 1, 2].DoubleValue; // column C (index 2)
            }

            if (totalSales > 300)
            {
                // Apply a dark built‑in style
                slicer.StyleType = SlicerStyleType.SlicerStyleDark2;
            }
            else
            {
                // Apply a light built‑in style
                slicer.StyleType = SlicerStyleType.SlicerStyleLight1;
            }

            // Optional: set additional slicer appearance
            slicer.Caption = "Product Filter";
            slicer.NumberOfColumns = 1;
            slicer.WidthPixel = 200;
            slicer.HeightPixel = 120;

            // Save the workbook
            workbook.Save("SlicerStyleDynamicDemo.xlsx");
        }
    }
}
