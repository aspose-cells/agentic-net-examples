using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsQuotePrefixDemo
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

                // Populate some numeric and non‑numeric values for demonstration
                cells["A1"].PutValue(12345);          // numeric
                cells["A2"].PutValue("67890");        // string that looks numeric
                cells["A3"].PutValue(12.34);          // numeric (double)
                cells["A4"].PutValue("Text");         // non‑numeric

                // Iterate through the used range of the worksheet
                AsposeRange usedRange = cells.MaxDisplayRange;
                int lastRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int lastCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = usedRange.FirstRow; row <= lastRow; row++)
                {
                    for (int col = usedRange.FirstColumn; col <= lastCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Apply QuotePrefix only to numeric cells
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            Style style = workbook.CreateStyle();
                            style.QuotePrefix = true;

                            StyleFlag flag = new StyleFlag();
                            flag.QuotePrefix = true;

                            cell.SetStyle(style, flag);
                        }
                    }
                }

                // Save the workbook – numeric cells will now be stored as text with a leading apostrophe
                string outputPath = "QuotePrefixDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}