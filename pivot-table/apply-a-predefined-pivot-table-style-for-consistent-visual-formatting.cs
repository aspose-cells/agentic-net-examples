// Title: Apply the built‑in PivotStyleLight16 to a pivot table using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, fills it with sample sales data, adds a pivot table, assigns PivotTableStyleName = "PivotStyleLight16", and saves the file as XLSX with Aspose.Cells. | Update the sample to use the built‑in style "PivotStyleMedium9" instead, and describe how to switch between predefined pivot styles programmatically. | Create a reusable method that receives a PivotTable object and a style name string, applies the style via PivotTableStyleName, and returns the modified workbook.
// Common Searches: how to change pivot table style with Aspose.Cells C# | Aspose.Cells set built‑in pivot style PivotStyleLight16 example | C# code to apply predefined pivot style to Excel workbook using Aspose.Cells | apply custom pivot table formatting Aspose.Cells .NET tutorial | Aspose.Cells PivotTableStyleName property usage guide
// Tags: Aspose.Cells pivot table style application | C# PivotTableStyleName usage | built‑in PivotStyleLight16 formatting | Excel pivot visual theme Aspose.Cells | predefined pivot table style .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotStyleDemo
{
    // Demonstrates creating a workbook, populating sample sales data, adding a pivot table, applying the built‑in "PivotStyleLight16" via the PivotTableStyleName property, and saving the result as an XLSX file using Aspose.Cells for .NET.
    public class ApplyPredefinedPivotStyle
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Year";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Electronics";
                sheet.Cells["B2"].Value = 2020;
                sheet.Cells["C2"].Value = 15000;

                sheet.Cells["A3"].Value = "Electronics";
                sheet.Cells["B3"].Value = 2021;
                sheet.Cells["C3"].Value = 18000;

                sheet.Cells["A4"].Value = "Furniture";
                sheet.Cells["B4"].Value = 2020;
                sheet.Cells["C4"].Value = 12000;

                sheet.Cells["A5"].Value = "Furniture";
                sheet.Cells["B5"].Value = 2021;
                sheet.Cells["C5"].Value = 14000;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("=Sheet1!A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Year as column field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as data field

                // Apply a predefined built‑in pivot table style
                pivotTable.PivotTableStyleName = "PivotStyleLight16";

                // Save the workbook
                string outputPath = "PivotTableWithPredefinedStyle.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyPredefinedPivotStyle.Run();
        }
    }
}
