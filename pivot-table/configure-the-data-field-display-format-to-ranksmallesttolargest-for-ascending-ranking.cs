using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotFieldRankSmallestToLargestDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data for the pivot table
                cells["A1"].Value = "Category";
                cells["B1"].Value = "Amount";
                cells["A2"].Value = "A";
                cells["B2"].Value = 30;
                cells["A3"].Value = "B";
                cells["B3"].Value = 20;
                cells["A4"].Value = "C";
                cells["B4"].Value = 50;
                cells["A5"].Value = "A";
                cells["B5"].Value = 10;
                cells["A6"].Value = "B";
                cells["B6"].Value = 40;
                cells["A7"].Value = "C";
                cells["B7"].Value = 60;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B7", "D3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add row field (Category) and data field (Amount)
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Set the data field to display rank from smallest to largest
                PivotField dataField = pivot.DataFields[0];
                dataField.ShowValuesSetting.CalculationType = PivotFieldDataDisplayFormat.RankSmallestToLargest;

                // Refresh and calculate the pivot table
                pivot.RefreshData();
                pivot.CalculateData();

                // Save the workbook
                string outputPath = "PivotField_RankSmallestToLargest.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
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
            PivotFieldRankSmallestToLargestDemo.Run();
        }
    }
}