using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemPositionDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Orange");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["B4"].PutValue(1500);

                // Add a pivot table based on the data range
                int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[ptIndex];

                // Add the "Product" field to the row area and "Sales" to the data area
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh the pivot table data and calculate the results
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Set the Position of each pivot item.
                // Use a for‑loop to avoid modifying the collection while enumerating it.
                PivotField rowField = pivotTable.RowFields[0];
                for (int i = 0; i < rowField.PivotItems.Count; i++)
                {
                    // Example: move each item to the first position (index 0)
                    rowField.PivotItems[i].Position = 0;
                }

                // Define output file path
                string outputPath = "PivotItemPositionAfterRefresh.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}