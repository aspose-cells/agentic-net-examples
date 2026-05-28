using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPivotExample
{
    public class LinkedPivotTableDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the default worksheet and add sample source data
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Populate source data (example table)
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["B1"].PutValue("Product");
                sourceSheet.Cells["C1"].PutValue("Sales");

                sourceSheet.Cells["A2"].PutValue("Fruit");
                sourceSheet.Cells["B2"].PutValue("Apple");
                sourceSheet.Cells["C2"].PutValue(1200);

                sourceSheet.Cells["A3"].PutValue("Fruit");
                sourceSheet.Cells["B3"].PutValue("Banana");
                sourceSheet.Cells["C3"].PutValue(850);

                sourceSheet.Cells["A4"].PutValue("Vegetable");
                sourceSheet.Cells["B4"].PutValue("Carrot");
                sourceSheet.Cells["C4"].PutValue(640);

                // Add a new worksheet that will host the linked PivotTable
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

                // Copy the source data to the new worksheet (optional, for visual reference)
                int lastRow = sourceSheet.Cells.MaxDataRow;
                int lastColumn = sourceSheet.Cells.MaxDataColumn;
                for (int row = 0; row <= lastRow; row++)
                {
                    for (int col = 0; col <= lastColumn; col++)
                    {
                        Cell srcCell = sourceSheet.Cells[row, col];
                        Cell destCell = pivotSheet.Cells[row, col];
                        destCell.PutValue(srcCell.Value);
                    }
                }

                // Build the source data reference string for the PivotTable cache
                AsposeRange srcRange = sourceSheet.Cells.MaxDisplayRange;
                string sourceData = $"={sourceSheet.Name}!{srcRange.Address}";

                // Add a linked PivotTable on the new worksheet using the source data reference
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceData, "A1", "LinkedPivotTable");

                // Configure the PivotTable (example: Category as row, Sales as data)
                PivotTable pivotTable = pivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the PivotTable data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "LinkedPivotDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    internal class Program
    {
        private static void Main(string[] args)
        {
            LinkedPivotTableDemo.Run();
        }
    }
}