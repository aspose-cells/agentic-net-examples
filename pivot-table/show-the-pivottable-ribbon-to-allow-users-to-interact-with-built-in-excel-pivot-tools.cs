using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotDemo
{
    class ShowPivotTableRibbonDemo
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
                sheet.Cells["B2"].Value = 100;
                sheet.Cells["A3"].Value = "Drink";
                sheet.Cells["B3"].Value = 150;
                sheet.Cells["A4"].Value = "Food";
                sheet.Cells["B4"].Value = 200;

                // Add a pivot table to the worksheet
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Enable UI features that expose the PivotTable ribbon/tools
                pivotTable.EnableFieldList = true;    // Shows the field list pane
                pivotTable.EnableFieldDialog = true; // Enables the field dialog on double‑click
                pivotTable.EnableWizard = true;      // Makes the PivotTable Wizard available

                // Refresh and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTableRibbonDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ShowPivotTableRibbonDemo.Run();
        }
    }
}