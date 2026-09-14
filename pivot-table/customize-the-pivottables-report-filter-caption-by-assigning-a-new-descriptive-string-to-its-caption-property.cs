// Title: How to set a custom caption for a PivotTable report filter (page field) using Aspose.Cells for .NET
// AI Prompts: Generate C# code that assigns a new display name to a PivotTable page field using Aspose.Cells. | Show how to update the report filter caption of a PivotTable and then refresh and recalculate it in a .NET workbook. | Create a full example that builds a PivotTable, changes the filter label to a custom text, and saves the file with Aspose.Cells.
// Common Searches: Aspose.Cells how to rename pivot table page field label in C# | set custom caption for pivot table report filter programmatically | change pivot table filter name using Aspose.Cells .NET | example of modifying pivot table filter caption in Aspose.Cells
// Tags: Aspose.Cells set pivot page field name | custom report filter caption .NET | rename pivot table filter label Aspose.Cells | pivot table caption property C# | refresh calculate pivot after label update Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a PivotTable, adds a page field for "Category", sets its Name property to "Select Category" to customize the filter caption, refreshes and calculates the PivotTable, and saves the workbook as PivotTableReportFilterCaptionDemo.xlsx.
    public class PivotTableReportFilterCaptionDemo
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
                sheet.Cells["A3"].Value = "Drink";
                sheet.Cells["B3"].Value = 80;
                sheet.Cells["A4"].Value = "Food";
                sheet.Cells["B4"].Value = 150;
                sheet.Cells["A5"].Value = "Drink";
                sheet.Cells["B5"].Value = 70;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");   // Data field

                // Add a report filter (page field) for the same "Category" field
                pivotTable.AddFieldToArea(PivotFieldType.Page, "Category");

                // Customize the report filter caption by setting a new name for the page field
                // The Name property of a page field acts as the caption displayed in the filter area
                pivotTable.PageFields[0].Name = "Select Category";

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();   // Correct API usage
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTableReportFilterCaptionDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
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
            PivotTableReportFilterCaptionDemo.Run();
        }
    }
}
