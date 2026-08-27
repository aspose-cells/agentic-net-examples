// Title: Auto‑fit multiple columns then set exact pixel widths with SetColumnWidthPixel in Aspose.Cells for .NET
// AI Prompts: Create a C# workbook, populate columns A‑E with sample data, call AutoFitColumns, and then set each column's width to a specific pixel value using Cells.SetColumnWidthPixel. | Write C# code that demonstrates how to auto‑fit a range of columns and subsequently override their widths with precise pixel measurements in an Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells C# set column width in pixels after AutoFitColumns for a range of columns | How to override auto‑fitted column widths with exact pixel sizes using Aspose.Cells .NET | C# example of applying SetColumnWidthPixel to multiple columns after auto‑fit in Excel workbook
// Tags: apply pixel widths after column auto‑fit Aspose.Cells | SetColumnWidthPixel for batch columns .NET | override column widths with exact pixels Aspose | precise Excel column sizing C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, fills columns A‑E with sample text, auto‑fits those columns, then overwrites each column's width with predefined pixel values using Cells.SetColumnWidthPixel, and finally saves the file as SetColumnWidthPixelAfterAutoFit.xlsx.
    public class SetColumnWidthPixelAfterAutoFit
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill columns A to E with sample data to demonstrate auto‑fit
            for (int col = 0; col < 5; col++)
            {
                for (int row = 0; row < 10; row++)
                {
                    cells[row, col].PutValue($"Row{row + 1} Col{col + 1} - Sample text");
                }
            }

            // Auto‑fit columns 0 through 4 based on the populated data
            sheet.AutoFitColumns(0, 4);

            // Define the exact pixel widths you want for each column after auto‑fit
            int[] targetPixelWidths = new int[] { 150, 120, 180, 130, 160 };

            // Apply the precise pixel widths using SetColumnWidthPixel
            for (int i = 0; i < targetPixelWidths.Length; i++)
            {
                cells.SetColumnWidthPixel(i, targetPixelWidths[i]);
            }

            // Determine output file path
            string outputPath = "SetColumnWidthPixelAfterAutoFit.xlsx";

            // Save the workbook with the adjusted column widths
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
    }
}
