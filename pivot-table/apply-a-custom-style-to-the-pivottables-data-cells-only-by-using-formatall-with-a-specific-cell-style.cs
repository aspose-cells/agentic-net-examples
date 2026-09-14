// Title: Apply a custom Calibri bold style with light‑yellow background to PivotTable data cells using FormatAll and PreserveFormatting in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a PivotTable, defines a bold Calibri style with a light‑yellow fill, sets PreserveFormatting to true, and applies the style to the data area using the FormatAll method in Aspose.Cells. | Show how to keep a custom cell style on PivotTable data cells after refreshing the table by calling FormatAll on the PivotTable object with a predefined Style in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells how to style only the values in a PivotTable using C# | Keep PivotTable formatting after data refresh with Aspose.Cells .NET | Example of applying a yellow background to PivotTable data cells in a generated XLSX | C# code to retain formatting and style PivotTable data area in Aspose.Cells | Apply bold Calibri formatting to PivotTable rows with Aspose.Cells library
// Tags: PivotTable cell formatting Aspose.Cells | PreserveFormatting property Aspose.Cells | C# define Calibri bold style Aspose.Cells | Apply style to pivot table data area XLSX | Aspose.Cells pivot table styling example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace PivotTableCustomStyleDemo
{
    // The sample creates a workbook, adds sample data, builds a PivotTable, defines a bold Calibri style with a light‑yellow background, enables PreserveFormatting, applies the style to all PivotTable cells via FormatAll, and saves the result as an XLSX file, ensuring the custom formatting persists after any refresh.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].Value = "Category";
            worksheet.Cells["B1"].Value = "Amount";
            worksheet.Cells["A2"].Value = "Food";
            worksheet.Cells["B2"].Value = 1200;
            worksheet.Cells["A3"].Value = "Food";
            worksheet.Cells["B3"].Value = 800;
            worksheet.Cells["A4"].Value = "Drink";
            worksheet.Cells["B4"].Value = 500;
            worksheet.Cells["A5"].Value = "Drink";
            worksheet.Cells["B5"].Value = 700;

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Category, data = Amount
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Ensure formatting is preserved when the pivot table recalculates
            pivotTable.PreserveFormatting = true;

            // Create a custom style for the data cells
            Style dataStyle = workbook.CreateStyle();
            dataStyle.Font.Name = "Calibri";
            dataStyle.Font.Size = 11;
            dataStyle.Font.IsBold = true;
            dataStyle.ForegroundColor = Color.LightYellow;
            dataStyle.Pattern = BackgroundType.Solid;

            // Apply the custom style to all cells in the pivot table area
            // (FormatAll formats the entire pivot table; with PreserveFormatting set,
            //  the style will be retained for data cells after any refresh)
            pivotTable.FormatAll(dataStyle);

            // Save the workbook
            workbook.Save("PivotTableDataStyleDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
