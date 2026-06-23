using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWrapTextDemo
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
                Cells cells = worksheet.Cells;

                // Fill range A1:C5 with a long text
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 3; col++)
                    {
                        cells[row, col].PutValue("This is a long text that should wrap inside the cell when wrapping is enabled.");
                    }
                }

                // Create a style with text wrapping enabled
                Style wrapStyle = workbook.CreateStyle();
                wrapStyle.IsTextWrapped = true;

                // Apply only the WrapText property via a StyleFlag
                StyleFlag flag = new StyleFlag();
                flag.WrapText = true;

                // Define the range A1:C5 (row 0, column 0, 5 rows, 3 columns)
                Aspose.Cells.Range range = cells.CreateRange(0, 0, 5, 3);

                // Apply the wrap style to the range
                range.ApplyStyle(wrapStyle, flag);

                // Auto‑fit rows so the wrapped text becomes visible
                worksheet.AutoFitRows();

                // Determine output file path
                string outputPath = "WrapTextRangeDemo.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}