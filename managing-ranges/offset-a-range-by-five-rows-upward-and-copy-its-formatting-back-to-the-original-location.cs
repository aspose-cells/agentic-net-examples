using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsOffsetFormattingDemo
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

                // Define the original range (e.g., A6:D10)
                AsposeRange originalRange = sheet.Cells.CreateRange("A6:D10");

                // Apply sample formatting to the original range
                Style style = workbook.CreateStyle();
                style.Font.Name = "Arial";
                style.Font.Size = 12;
                style.Font.IsBold = true;
                style.ForegroundColor = System.Drawing.Color.LightGreen;
                style.Pattern = BackgroundType.Solid;
                originalRange.SetStyle(style);

                // Fill the original range with sample values
                for (int row = 0; row < originalRange.RowCount; row++)
                {
                    for (int col = 0; col < originalRange.ColumnCount; col++)
                    {
                        originalRange[row, col].PutValue($"R{row + 6}C{col + 1}");
                    }
                }

                // Define the destination range that is 5 rows upward (A1:D5)
                AsposeRange offsetRange = sheet.Cells.CreateRange("A1:D5");

                // Move the original range content up by 5 rows
                originalRange.MoveTo(offsetRange.FirstRow, offsetRange.FirstColumn);

                // Copy the formatting from the moved range back to the original location
                originalRange.CopyStyle(offsetRange);

                // Save the workbook (ensure the directory exists)
                string outputPath = "OffsetCopyFormatting.xlsx";
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