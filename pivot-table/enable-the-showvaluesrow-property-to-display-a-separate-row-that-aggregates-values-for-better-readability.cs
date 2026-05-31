using System;
using System.IO;
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
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Amount";
                cells["A2"].Value = "Food";
                cells["B2"].Value = 120;
                cells["A3"].Value = "Food";
                cells["B3"].Value = 80;
                cells["A4"].Value = "Beverage";
                cells["B4"].Value = 150;
                cells["A5"].Value = "Beverage";
                cells["B5"].Value = 70;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: Category as row field, Amount as data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Category
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Amount

                // Enable the separate values row for better readability
                pivotTable.ShowValuesRow = true;

                // Save the workbook to a file
                string outputPath = "PivotTableShowValuesRowDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
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