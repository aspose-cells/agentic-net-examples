using System;
using Aspose.Cells;

namespace AsposeCellsReadAndPrint
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "input.xlsx";

            using (Workbook workbook = new Workbook(filePath))
            {
                for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
                {
                    Worksheet sheet = workbook.Worksheets[sheetIndex];
                    Console.WriteLine($"--- Worksheet [{sheetIndex}]: {sheet.Name} ---");

                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];
                            if (cell.Value == null) continue;

                            string address = cell.Name;
                            string value = cell.Value?.ToString() ?? string.Empty;

                            Style style = cell.GetStyle();
                            string fontName = style.Font.Name;
                            double fontSize = style.Font.Size;
                            string fontColor = style.Font.Color.ToString();

                            Console.WriteLine($"{address}: \"{value}\" [Font: {fontName}, Size: {fontSize}, Color: {fontColor}]");
                        }
                    }

                    Console.WriteLine();
                }
            }
        }
    }
}