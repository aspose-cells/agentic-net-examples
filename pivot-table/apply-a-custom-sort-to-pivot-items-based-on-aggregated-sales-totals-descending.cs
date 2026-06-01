using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCustomSort
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
                Cells cells = sheet.Cells;

                // Populate sample data: Product | Sales
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Sales";

                cells["A2"].Value = "Apple";
                cells["B2"].Value = 1200;

                cells["A3"].Value = "Banana";
                cells["B3"].Value = 800;

                cells["A4"].Value = "Cherry";
                cells["B4"].Value = 1500;

                cells["A5"].Value = "Date";
                cells["B5"].Value = 600;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add the Product field to the row area
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

                // Add the Sales field to the data area
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Retrieve the row field (Product) to apply custom sorting
                PivotField productField = pivotTable.RowFields[0];

                // Sort the Product items by the aggregated Sales totals in descending order.
                // Use the overload that sorts by data field index (0 = first data field).
                productField.SortBy(SortOrder.Descending, 0);

                // Refresh and calculate the pivot table to apply sorting
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Define output file
                string outputPath = "PivotCustomSortDescending.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

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
}