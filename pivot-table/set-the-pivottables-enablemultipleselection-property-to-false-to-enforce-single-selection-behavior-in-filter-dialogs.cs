// Title: How to disable multiple filter selections in an Aspose.Cells PivotTable (C#) by setting AllowMultipleFiltersPerField to false
// AI Prompts: Generate C# code that builds a workbook, adds sample data, creates a PivotTable with Aspose.Cells, and configures it to allow only one filter choice per field. | Provide the sequence to refresh, calculate, and save the workbook after restricting the PivotTable to a single‑selection mode.
// Common Searches: Aspose.Cells C# pivot table single selection filter dialog | set AllowMultipleFiltersPerField false Aspose.Cells example | disable multiple filters per field in Aspose.Cells PivotTable | how to enforce single selection in pivot table filters using .NET | Aspose.Cells pivot table filter settings for single choice
// Tags: Aspose.Cells pivot filter exclusive mode | configure pivot table filter behavior Aspose.Cells | create pivot table Aspose.Cells C# | save workbook as xlsx Aspose.Cells | pivot filter multiple selection toggle

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable and PivotFieldType

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, populating sample data, adding a PivotTable with Aspose.Cells for .NET, disabling multiple filter selections by setting AllowMultipleFiltersPerField to false, refreshing and calculating the PivotTable, and saving the result as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Product";
                sheet.Cells["B1"].Value = "Category";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Apple";
                sheet.Cells["B2"].Value = "Fruit";
                sheet.Cells["C2"].Value = 1000;

                sheet.Cells["A3"].Value = "Banana";
                sheet.Cells["B3"].Value = "Fruit";
                sheet.Cells["C3"].Value = 1500;

                sheet.Cells["A4"].Value = "Carrot";
                sheet.Cells["B4"].Value = "Vegetable";
                sheet.Cells["C4"].Value = 800;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Enforce single‑selection behavior in filter dialogs
                pivotTable.AllowMultipleFiltersPerField = false;

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Define output file path
                string outputPath = "PivotTable_SingleSelection.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the pivot table workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
