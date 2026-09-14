// Title: How to apply a LightBlue fill to a pivot table header cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a solid LightBlue style and assigns it to the row header of a PivotTable via the PivotTable.Format method. | Demonstrate how to style both row and column headers of an Aspose.Cells pivot table with a LightBlue background.
// Common Searches: Aspose.Cells C# set background color of pivot table row header | apply solid fill to pivot table header cell using Aspose.Cells .NET | how to change pivot field header color in Excel workbook with Aspose.Cells | C# example for formatting pivot table header background color Aspose.Cells
// Tags: pivot table header background color Aspose.Cells | C# PivotTable.Format style application | Aspose.Cells solid fill style for Excel header | lightblue cell style in Aspose.Cells workbook | formatting pivot field header cell .NET

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a pivot table, defines a solid LightBlue style, applies it to the pivot table's row header (and optionally column header) using PivotTable.Format, and saves the result as an .xlsx file.
    public class ApplyLightBlueFillToPivotHeader
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
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: Category as row field, Amount as data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column A (Category)
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Column B (Amount)

                // Refresh and calculate the pivot table so that headers are generated
                pivotTable.RefreshData();   // Correct API call
                pivotTable.CalculateData();

                // Create a style with LightBlue fill color
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Pattern = BackgroundType.Solid;
                headerStyle.ForegroundColor = Color.LightBlue;

                // Apply the style to the pivot table header cell (row header)
                pivotTable.Format(0, 0, headerStyle);
                // To style the column header (if present), uncomment the following line:
                // pivotTable.Format(0, 1, headerStyle);

                // Save the workbook
                string outputPath = "PivotHeaderLightBlue.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyLightBlueFillToPivotHeader.Run();
        }
    }
}
