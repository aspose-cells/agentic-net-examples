using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableReportFilterCaptionDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 600;

            sheet.Cells["A5"].Value = "Vegetable";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 900;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");    // Data field

            // Add a report filter (page field)
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Region");

            // Customize the caption of the report filter by setting a new name for the page field
            // This name is displayed as the filter caption in the PivotTable UI
            pivotTable.PageFields[0].Name = "Sales Region Filter";

            // Optionally, display the filter pages for the customized caption
            pivotTable.ShowReportFilterPageByName("Sales Region Filter");

            // Refresh and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Determine output path and save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PivotTableReportFilterCaptionDemo.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}