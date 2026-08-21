// Title: C# – Apply a custom style to PivotTable data cells with FormatAll and PreserveFormatting in Aspose.Cells
// Description: Creates a workbook, adds sample sales data, builds a PivotTable, defines a Calibri bold style with a light‑yellow background, enables PreserveFormatting, and calls FormatAll so only the data cells receive the custom formatting. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# PivotTable style | FormatAll | PreserveFormatting | custom cell style | Excel PivotTable formatting | .NET Excel library | pivot data cell formatting | XLSX report styling | global developers
// Common Searches: Aspose.Cells format only pivot table data cells | C# FormatAll PreserveFormatting PivotTable | apply custom background to pivot values Aspose | style pivot table data area .NET | Aspose.Cells custom style for pivot table values
// Developer Intent: Apply a custom cell style exclusively to the data area of a PivotTable using Aspose.Cells for .NET.
// Use Cases: Generate a sales report where summed amounts are highlighted with bold Calibri text on a light‑yellow background while row headers stay unchanged. | Create a financial workbook that emphasizes aggregated figures by styling only the PivotTable data cells, preserving default header formatting. | Export an Excel file with a PivotTable where data cells have consistent styling for clearer downstream analysis.
// AI Prompts: Show how to style only the data cells of a PivotTable in Aspose.Cells using FormatAll with PreserveFormatting. | Provide a C# example that creates a custom style, sets PreserveFormatting, and applies it to PivotTable values. | Explain why PreserveFormatting must be true when using FormatAll to affect only the data area of a PivotTable.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableCustomStyleDemo
{
    // Creates a workbook, adds sample sales data, builds a PivotTable, defines a Calibri bold style with a light‑yellow background, enables PreserveFormatting, and calls FormatAll so only the data cells receive the custom formatting. The workbook is saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = "Food";
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["A4"].Value = "Drink";
            sheet.Cells["B4"].Value = 150;
            sheet.Cells["A5"].Value = "Drink";
            sheet.Cells["B5"].Value = 70;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Category, data = Sum of Amount
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Create a custom style for the data cells
            Style dataStyle = workbook.CreateStyle();
            dataStyle.Font.Name = "Calibri";
            dataStyle.Font.Size = 11;
            dataStyle.Font.IsBold = true;
            dataStyle.ForegroundColor = Color.LightYellow;
            dataStyle.Pattern = BackgroundType.Solid;

            // Apply the custom style to all cells in the pivot table area
            // (FormatAll formats the entire pivot table; to affect only data cells,
            //  ensure PreserveFormatting is true so other areas retain their default formatting)
            pivotTable.PreserveFormatting = true;
            pivotTable.FormatAll(dataStyle);

            // Save the workbook
            workbook.Save("PivotTableCustomStyle.xlsx", SaveFormat.Xlsx);
        }
    }
}
