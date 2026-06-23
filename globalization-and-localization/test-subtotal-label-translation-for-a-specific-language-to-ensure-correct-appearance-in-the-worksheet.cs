using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSubtotalTranslationDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table.
                // Column A – Category, Column B – Value.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("A");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("B");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("B");
                sheet.Cells["B5"].PutValue(40);

                // Create a pivot table that starts at cell D1.
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "DemoPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add the Category field to the Row area and the Value field to the Data area.
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column A (Category)
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Column B (Value)

                // Refresh and calculate the pivot table so that it reflects the data.
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook.
                string outputPath = "PivotSubtotalTranslationDemo.xlsx";

                // Ensure the directory exists before saving.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}