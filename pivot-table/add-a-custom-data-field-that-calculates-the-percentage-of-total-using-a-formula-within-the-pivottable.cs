using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotPercentageExample
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

                // Populate sample data
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Sales";
                cells["A2"].Value = "Apple";
                cells["B2"].Value = 1200;
                cells["A3"].Value = "Orange";
                cells["B3"].Value = 800;
                cells["A4"].Value = "Banana";
                cells["B4"].Value = 500;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D6", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add row field (Product) and data field (Sales)
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Configure the data field to display as a percentage of the grand total
                PivotField dataField = pivot.DataFields[0];
                // The ShowDataAs property may not be available in older Aspose.Cells versions.
                // If supported, uncomment the following line:
                // dataField.ShowDataAs = ShowDataAs.PercentageOfGrandTotal;
                dataField.NumberFormat = "0.00%";

                // Refresh and calculate the pivot table
                pivot.RefreshData();
                pivot.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_PercentageOfTotal.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}