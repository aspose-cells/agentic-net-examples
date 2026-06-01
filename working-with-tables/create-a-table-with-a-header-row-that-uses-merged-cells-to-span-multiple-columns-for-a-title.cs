using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class MergedHeaderTableDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // 1. Create a title that spans across 4 columns (A1:D1)
            cells.Merge(0, 0, 1, 4);
            cells[0, 0].PutValue("Sales Report 2024");
            Style titleStyle = cells[0, 0].GetStyle();
            titleStyle.HorizontalAlignment = TextAlignmentType.Center;
            titleStyle.Font.IsBold = true;
            titleStyle.Font.Size = 14;
            cells[0, 0].SetStyle(titleStyle);

            // 2. Add column headers (row 2)
            string[] headers = { "Region", "Product", "Quantity", "Revenue" };
            for (int col = 0; col < headers.Length; col++)
            {
                cells[1, col].PutValue(headers[col]);
                Style headerStyle = cells[1, col].GetStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.HorizontalAlignment = TextAlignmentType.Center;
                cells[1, col].SetStyle(headerStyle);
            }

            // 3. Add sample data starting from row 3
            object[,] data = {
                { "North", "Apples", 120, 2400 },
                { "South", "Oranges", 85, 1700 },
                { "East", "Bananas", 150, 3000 },
                { "West", "Grapes", 60, 1800 }
            };

            for (int row = 0; row < data.GetLength(0); row++)
            {
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    cells[2 + row, col].PutValue(data[row, col]);
                }
            }

            // 4. Auto‑fit columns to display content nicely
            worksheet.AutoFitColumns();

            // Save the workbook to a file
            string outputPath = "MergedHeaderTable.xlsx";
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                Console.Error.WriteLine($"Failed to save workbook: {saveEx.Message}");
                throw;
            }
        }
    }
}