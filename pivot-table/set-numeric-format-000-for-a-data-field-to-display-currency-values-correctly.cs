using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class SetNumericFormatDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Sales");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(1234.56);
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["B3"].PutValue(2345.78);
                worksheet.Cells["A4"].PutValue("Banana");
                worksheet.Cells["B4"].PutValue(3456.90);

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add the product field to the row area
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

                // Add the sales field to the data area
                int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                PivotField dataField = pivotTable.DataFields[dataFieldPos];

                // Set aggregation function to Sum
                dataField.Function = ConsolidationFunction.Sum;

                // Apply custom numeric format '#,##0.00'
                dataField.NumberFormat = "#,##0.00";

                // Refresh and calculate the pivot table to apply the format
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "SetNumericFormatDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            SetNumericFormatDemo.Run();
        }
    }
}