using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class SetRankLargestToSmallestDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Amount";
                cells["A2"].Value = "A";
                cells["B2"].Value = 10;
                cells["A3"].Value = "B";
                cells["B3"].Value = 30;
                cells["A4"].Value = "C";
                cells["B4"].Value = 20;
                cells["A5"].Value = "A";
                cells["B5"].Value = 15;
                cells["A6"].Value = "B";
                cells["B6"].Value = 25;
                cells["A7"].Value = "C";
                cells["B7"].Value = 5;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B7", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add a row field (Category)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

                // Add a data field (Amount)
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Set the data field to display rank from largest to smallest
                PivotField dataField = pivotTable.DataFields[0];
                dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.RankLargestToSmallest;

                // Calculate the pivot data
                pivotTable.CalculateData();

                // Define output file path
                string outputPath = "SetRankLargestToSmallestDemo_out.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetRankLargestToSmallestDemo.Run();
        }
    }
}