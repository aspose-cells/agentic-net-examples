using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
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

                // Populate sample data for the table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.2);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.8);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(2.5);

                // Add a ListObject (Excel table) that includes the header row
                int tableIdx = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIdx];
                table.TableStyleType = TableStyleType.TableStyleLight1; // optional visual style

                // Create a style with a light gray fill pattern
                Style style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Gray25;          // 25% gray pattern (light gray)
                style.ForegroundColor = Color.LightGray;       // Light gray foreground
                style.BackgroundColor = Color.White;           // White background

                // Enable the cell shading flag so the pattern is applied
                StyleFlag flag = new StyleFlag { CellShading = true };

                // Apply the style to the data range of the table (includes header if needed)
                table.DataRange.ApplyStyle(style, flag);

                // Save the workbook
                string outputPath = "TableLightGrayFill.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}