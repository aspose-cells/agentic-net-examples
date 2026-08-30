// Title: Apply custom cell style, preserve formatting, and save a modified Aspose.Cells pivot table as XLSX in C#
// AI Prompts: Create a workbook, add a pivot table, set SaveData = false, enable PreserveFormatting, assign PivotStyleMedium9, format the data body range with a custom yellow background style, and save the file as ModifiedPivot.xlsx using Aspose.Cells for .NET. | Update an existing pivot table to stop embedding source data, keep its layout after refresh, apply both a built‑in pivot style and a custom Calibri bold style to the data area, then export the workbook to XLSX format with Aspose.Cells C#.
// Common Searches: Aspose.Cells C# how to prevent embedding source data in a pivot table | C# preserve pivot table formatting after refresh with Aspose.Cells | Apply custom background color to pivot table data area using Aspose.Cells .NET | Export workbook containing styled pivot table to XLSX with Aspose.Cells | Set built‑in pivot style and custom cell style for a pivot table in C# Aspose.Cells
// Tags: pivot table custom cell style Aspose.Cells | preserve pivot formatting Aspose.Cells .NET | disable source data embedding pivot Aspose.Cells | save workbook as XLSX Aspose.Cells C# | apply built‑in pivot style Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSaveDemo
{
    // The example creates a workbook, populates sample data, adds a pivot table, disables source data embedding, preserves formatting on refresh, applies a built‑in PivotStyleMedium9 and a custom yellow background style to the data body range, and finally saves the workbook as ModifiedPivot.xlsx in XLSX format using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue("Banana");
            sheet.Cells["C3"].PutValue(800);

            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue("Carrot");
            sheet.Cells["C4"].PutValue(600);

            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue("Broccoli");
            sheet.Cells["C5"].PutValue(700);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ----- Apply pivot table changes -----
            // 1. Do not embed source data with the workbook
            pivotTable.SaveData = false;

            // 2. Preserve formatting when the pivot table is refreshed
            pivotTable.PreserveFormatting = true;

            // 3. Apply a built‑in pivot table style
            pivotTable.PivotTableStyleName = "PivotStyleMedium9";

            // 4. Apply a custom style to the data body range
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.Name = "Calibri";
            customStyle.Font.Size = 11;
            customStyle.Font.IsBold = true;
            customStyle.ForegroundColor = Color.LightYellow;
            customStyle.Pattern = BackgroundType.Solid;

            // Format the entire data body area of the pivot table
            CellArea dataArea = pivotTable.DataBodyRange;
            pivotTable.Format(dataArea, customStyle);

            // Save the modified workbook as XLSX (uses Workbook.Save(string, SaveFormat) rule)
            workbook.Save("ModifiedPivot.xlsx", SaveFormat.Xlsx);

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Workbook saved successfully as ModifiedPivot.xlsx");
        }
    }
}
