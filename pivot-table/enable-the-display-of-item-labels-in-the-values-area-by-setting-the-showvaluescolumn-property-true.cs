using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ShowValuesColumnDemo
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
                sheet.Cells["B1"].PutValue("SubCategory");
                sheet.Cells["C1"].PutValue("Amount");

                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(120);

                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["B3"].PutValue("Banana");
                sheet.Cells["C3"].PutValue(80);

                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B4"].PutValue("Carrot");
                sheet.Cells["C4"].PutValue(50);

                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue("Broccoli");
                sheet.Cells["C5"].PutValue(70);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category as row
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // SubCategory as column
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Amount as data

                // The ShowValuesColumn property is not available in the current Aspose.Cells version.
                // If needed, other display options can be set here.

                // Save the workbook
                workbook.Save("ShowValuesColumnDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ShowValuesColumnDemo.Run();
        }
    }
}