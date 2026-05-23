using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ApplyBoldToEntireColumns
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create a range that covers cells G3:H3
                // G = column index 6, H = column index 7, row 3 = index 2 (zero‑based)
                Aspose.Cells.Range range = cells.CreateRange("G3", "H3");

                // Get the entire columns that contain the range (columns G and H)
                Aspose.Cells.Range entireColumns = range.EntireColumn;

                // Define a style with bold font
                Style boldStyle = workbook.CreateStyle();
                boldStyle.Font.IsBold = true;

                // Specify that only the bold attribute should be applied
                StyleFlag flag = new StyleFlag();
                flag.FontBold = true;

                // Apply the bold style to the entire columns
                entireColumns.ApplyStyle(boldStyle, flag);

                // Save the workbook
                string outputPath = "EntireColumnBold.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main()
        {
            Run();
        }
    }
}