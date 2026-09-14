// Title: Create a Profit Margin calculated field in an Aspose.Cells pivot table with C# and format it as percentage
// AI Prompts: Create a new calculated field called ProfitMargin that divides the Profit column by the Revenue column in an Aspose.Cells pivot table (C#). | Apply a percentage number format to the ProfitMargin field and trigger a pivot table refresh to recalculate data using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add custom calculated field to pivot table | example of profit margin calculation in Aspose.Cells pivot table | set percentage format for pivot table calculated field Aspose.Cells | recalculate pivot table after adding calculated field with Aspose.Cells | programmatically create profit margin field in Excel pivot using Aspose.Cells .NET
// Tags: profit margin formula Aspose.Cells | pivot table percentage formatting C# | refresh pivot cache Aspose.Cells .NET | add profit margin to Excel pivot using Aspose.Cells | Aspose.Cells pivot table example C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The sample creates a workbook with product, revenue, and profit data, builds a pivot table, adds a calculated field named ProfitMargin that computes Profit divided by Revenue, formats this field as a percentage, refreshes and calculates the pivot data, and saves the workbook as PivotTable_With_ProfitMargin.xlsx.
    public class AddProfitMarginCalculatedField
    {
        public static void Main()
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Product, Revenue, Profit
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Revenue";
            cells["C1"].Value = "Profit";

            cells["A2"].Value = "A";
            cells["B2"].Value = 1000;
            cells["C2"].Value = 200;

            cells["A3"].Value = "B";
            cells["B3"].Value = 1500;
            cells["C3"].Value = 300;

            cells["A4"].Value = "C";
            cells["B4"].Value = 2000;
            cells["C4"].Value = 500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Profit");

            // Add a calculated field "ProfitMargin" = Profit / Revenue
            // The formula must start with '=' and reference the source field names
            pivotTable.AddCalculatedField("ProfitMargin", "=Profit/Revenue", true);

            // Optionally format the calculated field as a percentage
            PivotField profitMarginField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
            profitMarginField.NumberFormat = "0.00%";

            // Refresh the pivot cache and calculate data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Ensure output directory exists (handle possible null directory)
            string outputPath = "PivotTable_With_ProfitMargin.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
