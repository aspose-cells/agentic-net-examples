using System;
using System.Drawing;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsAIDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string outputFile = "SalesReport.xlsx";

            // Create a new workbook
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Sales Report";

            // Header
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["D1"].PutValue("Total");

            // Sample data
            var products = new[] { "Apple", "Banana", "Orange", "Grape", "Mango" };
            for (int i = 0; i < products.Length; i++)
            {
                int row = i + 2;
                sheet.Cells[row, 0].PutValue(products[i]);
                sheet.Cells[row, 1].PutValue((i + 1) * 10);      // Quantity
                sheet.Cells[row, 2].PutValue((i + 1) * 1.5);    // Price
                sheet.Cells[row, 3].Formula = $"=B{row}*C{row}"; // Total formula
            }

            // Apply header style
            var style = workbook.CreateStyle();
            style.Font.IsBold = true;
            style.ForegroundColor = Color.LightGray;
            style.Pattern = BackgroundType.Solid;

            var styleFlag = new StyleFlag { All = true };
            var headerRange = sheet.Cells.CreateRange("A1:D1");
            headerRange.ApplyStyle(style, styleFlag);

            // Auto-fit columns
            sheet.AutoFitColumns();

            // Save the workbook
            workbook.Save(outputFile);

            Console.WriteLine($"Spreadsheet generated and saved to: {outputFile}");
        }
    }
}