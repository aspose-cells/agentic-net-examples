using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add integer values to cells
                worksheet.Cells["A1"].PutValue(123456);
                worksheet.Cells["A2"].PutValue(789012);

                // Create a style that displays numbers as integers (no decimal places)
                Style integerStyle = workbook.CreateStyle();
                integerStyle.Number = 1; // Built‑in format "0"

                // Apply the integer style to the desired range
                Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:A2");
                StyleFlag flag = new StyleFlag { NumberFormat = true };
                range.ApplyStyle(integerStyle, flag);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Ensure the full formatted data is shown even if it exceeds column width
                    FormatDataIgnoreColumnWidth = true
                };

                // Save the workbook as HTML
                string outputPath = "Integers.html";
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}