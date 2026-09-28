// Title: Disable subtotals for the SubCategory row field in an Aspose.Cells pivot table using C#
// AI Prompts: Generate C# code with Aspose.Cells that builds a pivot table from a range and turns off automatic subtotals for a selected row field. | Show how to retrieve a PivotField in Aspose.Cells, set its IsAutoSubtotals property to false, then refresh, calculate, and save the workbook. | Provide a step‑by‑step C# example that adds row and data fields to a pivot table and disables subtotals for the SubCategory field.
// Common Searches: Aspose.Cells C# hide subtotals for a pivot row field | disable auto subtotals for SubCategory in Aspose.Cells pivot table | C# example to turn off subtotals in Aspose.Cells pivot table rows | how to set IsAutoSubtotals false for a specific field in Aspose.Cells | remove subtotal rows from Aspose.Cells pivot table using C#
// Tags: Aspose.Cells pivot row field subtotal suppression | C# disable pivot field auto subtotals | Aspose.Cells pivot refresh and calculate | C# create pivot table from data range | Aspose.Cells save workbook as xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHideSubtotals
{
    // The sample creates a new workbook, populates it with Category, SubCategory, and Sales data, adds a pivot table on range A1:C5, places Category and SubCategory as row fields and Sales as a data field, then disables automatic subtotals for the SubCategory row field by setting IsAutoSubtotals to false. After refreshing and recalculating the pivot, the workbook is saved as PivotHideSubtotalsDemo.xlsx.
    public class HideSubtotalsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                // Columns: Category, SubCategory, Sales
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "SubCategory";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Electronics";
                sheet.Cells["B2"].Value = "TV";
                sheet.Cells["C2"].Value = 1200;

                sheet.Cells["A3"].Value = "Electronics";
                sheet.Cells["B3"].Value = "Radio";
                sheet.Cells["C3"].Value = 300;

                sheet.Cells["A4"].Value = "Clothing";
                sheet.Cells["B4"].Value = "Shirt";
                sheet.Cells["C4"].Value = 500;

                sheet.Cells["A5"].Value = "Clothing";
                sheet.Cells["B5"].Value = "Pants";
                sheet.Cells["C5"].Value = 700;

                // Add a pivot table based on the data range A1:C5, place it at E3
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add two row fields: Category and SubCategory
                int categoryRowIndex = pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                int subCategoryRowIndex = pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");

                // Add the Sales field as a data field
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Hide subtotals for the "SubCategory" row field only
                PivotField subCategoryField = pivotTable.RowFields[subCategoryRowIndex];
                subCategoryField.IsAutoSubtotals = false;

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();      // Refreshes the pivot cache
                pivotTable.CalculateData();    // Recalculates the pivot table

                // Ensure the output directory exists
                string outputPath = "PivotHideSubtotalsDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
