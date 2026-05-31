using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDateImportDemo
{
    // Alias to avoid conflict with System.Range (C# 8+)
    using AsposeRange = Aspose.Cells.Range;

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a new worksheet
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet worksheet = workbook.Worksheets[newSheetIndex];

                // Prepare an array of DateTime objects
                object[] dateArray = new object[]
                {
                    new DateTime(2023, 1, 10),
                    new DateTime(2023, 2, 20),
                    new DateTime(2023, 3, 30)
                };

                // Import the dates vertically starting at cell A1 (row 0, column 0)
                worksheet.Cells.ImportObjectArray(dateArray, 0, 0, true);

                // Create a style with short date number format (built‑in format 14)
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Number = 14; // Short date pattern

                // Apply the style to the imported range (A1:A3)
                AsposeRange dateRange = worksheet.Cells.CreateRange(0, 0, dateArray.Length, 1);
                StyleFlag flag = new StyleFlag { All = true };
                dateRange.ApplyStyle(dateStyle, flag);

                // Define output file path
                string outputPath = "DateArrayImport.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}