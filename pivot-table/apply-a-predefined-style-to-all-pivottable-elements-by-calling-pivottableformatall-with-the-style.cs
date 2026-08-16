// Title: Apply a Custom Style to All PivotTable Elements with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data, build a PivotTable, define a Calibri bold style with a light‑gray background, and apply that style to every PivotTable cell using PivotTable.FormatAll before saving as XLSX.
// Keywords: Aspose.Cells PivotTable FormatAll | C# PivotTable styling | apply style to entire pivot table | Aspose.Cells custom style | PivotTable.FormatAll example | Aspose.Cells .NET formatting
// Common Searches: Aspose.Cells PivotTable.FormatAll C# example | how to style all cells of a pivot table using Aspose.Cells | apply custom formatting to Aspose.Cells pivot table | C# code to format entire pivot table Aspose | set background color for all pivot table cells Aspose.Cells
// Developer Intent: Programmatically apply a single custom style to every cell of a PivotTable in an Aspose.Cells workbook using the FormatAll method.
// Use Cases: Standardize the appearance of sales‑report pivot tables with corporate branding. | Quickly enforce consistent formatting across multiple generated pivot tables. | Apply a uniform style to a pivot table without configuring row, column, data, and page areas individually.
// AI Prompts: Generate C# code that creates a custom style and applies it to all elements of an Aspose.Cells PivotTable using PivotTable.FormatAll. | Show how to update the style of an already formatted PivotTable by calling FormatAll again with a new Style object. | Provide an example that uses one of Aspose.Cells built‑in styles (e.g., "PivotStyleMedium9") with FormatAll to style a pivot table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, build a PivotTable, define a Calibri bold style with a light‑gray background, and apply that style to every PivotTable cell using PivotTable.FormatAll before saving as XLSX.
    public class ApplyStyleToAllPivotTableElements
    {
        // Entry point for the application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].Value = "Category";
            worksheet.Cells["B1"].Value = "Year";
            worksheet.Cells["C1"].Value = "Amount";

            worksheet.Cells["A2"].Value = "Fruit";
            worksheet.Cells["B2"].Value = 2020;
            worksheet.Cells["C2"].Value = 150;

            worksheet.Cells["A3"].Value = "Fruit";
            worksheet.Cells["B3"].Value = 2021;
            worksheet.Cells["C3"].Value = 200;

            worksheet.Cells["A4"].Value = "Vegetable";
            worksheet.Cells["B4"].Value = 2020;
            worksheet.Cells["C4"].Value = 120;

            worksheet.Cells["A5"].Value = "Vegetable";
            worksheet.Cells["B5"].Value = 2021;
            worksheet.Cells["C5"].Value = 180;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = worksheet.PivotTables;
            int pivotIndex = pivotTables.Add("=Sheet1!A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Year");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Create a style that will be applied to the entire pivot table
            Style style = workbook.CreateStyle();
            style.Font.Name = "Calibri";
            style.Font.Size = 11;
            style.Font.IsBold = true;
            style.ForegroundColor = Color.LightGray;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to all cells of the pivot table
            pivotTable.FormatAll(style);

            // Save the workbook
            workbook.Save("PivotTable_Formatted_All.xlsx", SaveFormat.Xlsx);
        }
    }
}
