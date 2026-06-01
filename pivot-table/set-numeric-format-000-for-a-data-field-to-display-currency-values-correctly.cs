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
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1234.56);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(7890.12);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(3456.78);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the product field to the row area
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");

                // Add the sales field to the data area and obtain the PivotField object
                int dataFieldIndex = pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
                PivotField dataField = pivot.DataFields[dataFieldIndex];

                // Set custom numeric format (currency style)
                dataField.NumberFormat = "#,##0.00";

                // Refresh and calculate the pivot table to apply the format
                pivot.RefreshData();
                pivot.CalculateData();

                // Save the workbook
                string outputPath = "SetNumericFormatDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetNumericFormatDemo.Run();
        }
    }
}