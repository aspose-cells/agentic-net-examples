using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotTableShowValuesRowDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                Cells cells = sheet.Cells;
                cells["A1"].Value = "Fruit";
                cells["B1"].Value = "Quantity";
                cells["A2"].Value = "Apple";
                cells["B2"].Value = 10;
                cells["A3"].Value = "Orange";
                cells["B3"].Value = 15;
                cells["A4"].Value = "Banana";
                cells["B4"].Value = 20;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table (Fruit as row, Quantity as data)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

                // Hide the additional aggregated values row
                pivotTable.ShowValuesRow = false;

                // Save the workbook to a file (overwrite if it already exists)
                string outputPath = "PivotTableShowValuesRowDemo.xlsx";
                if (System.IO.File.Exists(outputPath))
                {
                    System.IO.File.Delete(outputPath);
                }
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableShowValuesRowDemo.Run();
        }
    }
}