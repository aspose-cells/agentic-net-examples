using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ConsolidatedPivotTableDemo
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
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ---------- Worksheet 1 ----------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                Cells cells1 = sheet1.Cells;
                // Header
                cells1["A1"].PutValue("Category");
                cells1["B1"].PutValue("Product");
                cells1["C1"].PutValue("Amount");
                // Sample rows
                cells1["A2"].PutValue("Food");
                cells1["B2"].PutValue("Apple");
                cells1["C2"].PutValue(120);
                cells1["A3"].PutValue("Food");
                cells1["B3"].PutValue("Banana");
                cells1["C3"].PutValue(80);
                cells1["A4"].PutValue("Drink");
                cells1["B4"].PutValue("Tea");
                cells1["C4"].PutValue(50);

                // ---------- Worksheet 2 ----------
                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                Cells cells2 = sheet2.Cells;
                // Header
                cells2["A1"].PutValue("Category");
                cells2["B1"].PutValue("Product");
                cells2["C1"].PutValue("Amount");
                // Sample rows
                cells2["A2"].PutValue("Food");
                cells2["B2"].PutValue("Apple");
                cells2["C2"].PutValue(150);
                cells2["A3"].PutValue("Food");
                cells2["B3"].PutValue("Orange");
                cells2["C3"].PutValue(90);
                cells2["A4"].PutValue("Drink");
                cells2["B4"].PutValue("Coffee");
                cells2["C4"].PutValue(70);

                // ---------- Pivot Table Worksheet ----------
                Worksheet pivotSheet = workbook.Worksheets.Add("ConsolidatedPivot");

                // Define multiple consolidation ranges (source data from both worksheets)
                string[] sourceRanges = { "=Sheet1!A1:C4", "=Sheet2!A1:C4" };

                // No page fields needed for this example
                PivotPageFields pageFields = new PivotPageFields();

                // Add the consolidated pivot table
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceRanges, false, pageFields, "A3", "ConsolidatedPivotTable");
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // The default aggregation for data fields is Sum, so no explicit function setting is required

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "ConsolidatedPivotTableDemo.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Runtime safety: capture any errors that occur during processing
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}