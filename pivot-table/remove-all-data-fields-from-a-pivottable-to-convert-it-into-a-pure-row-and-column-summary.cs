using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class RemoveAllDataFieldsFromPivotTable
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue(15);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add a row field and a data field (initially)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // ------------------------------------------------------------
                // Remove all data fields from the pivot table
                // ------------------------------------------------------------
                // Remove from last to first to avoid collection modification issues
                for (int i = pivotTable.DataFields.Count - 1; i >= 0; i--)
                {
                    string fieldName = pivotTable.DataFields[i].Name;
                    pivotTable.RemoveField(PivotFieldType.Data, fieldName);
                }

                // Refresh after removal
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Ensure output directory exists
                string outputPath = "PivotTable_NoDataFields.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveAllDataFieldsFromPivotTable.Run();
        }
    }
}