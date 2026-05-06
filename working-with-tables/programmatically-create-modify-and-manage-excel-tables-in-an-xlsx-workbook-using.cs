using System;
using Aspose.Cells;

namespace ExcelTableManagementDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Access the default worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 3. Define table headers
            string[] headers = { "Product", "Quantity", "UnitPrice" };
            for (int col = 0; col < headers.Length; col++)
            {
                cells[0, col].PutValue(headers[col]);
            }

            // 4. Populate sample data rows
            object[,] data = {
                { "Apple",  10, 0.5 },
                { "Banana", 20, 0.3 },
                { "Orange", 15, 0.4 }
            };

            for (int row = 0; row < data.GetLength(0); row++)
            {
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    cells[row + 1, col].PutValue(data[row, col]);
                }
            }

            // 5. Add a new column "Total" (Quantity * UnitPrice) using a formula
            int totalColIndex = headers.Length; // index after existing columns
            cells[0, totalColIndex].PutValue("Total"); // header

            // Apply formula for each data row
            for (int row = 1; row <= data.GetLength(0); row++)
            {
                // Formula: =B{row}*C{row}
                string formula = $"=B{row + 1}*C{row + 1}";
                cells[row, totalColIndex].Formula = formula;
            }

            // 6. Apply simple style to header row (bold font)
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            StyleFlag flag = new StyleFlag { All = true };
            cells.CreateRange("A1", "D1").ApplyStyle(headerStyle, flag);

            // 7. Adjust column widths for better visibility
            for (int col = 0; col <= totalColIndex; col++)
            {
                sheet.Cells.SetColumnWidth(col, 15);
            }

            // 8. Save the workbook to disk
            string outputPath = "ManagedTableDemo.xlsx";
            workbook.Save(outputPath);

            // Clean up
            workbook.Dispose();

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}