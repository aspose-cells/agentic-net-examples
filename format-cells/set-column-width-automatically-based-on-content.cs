using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class AutoFitColumnsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate cells with varying length text to demonstrate auto‑fit
                worksheet.Cells["A1"].PutValue("Short");
                worksheet.Cells["B1"].PutValue("Medium length text");
                worksheet.Cells["C1"].PutValue("This is a much longer piece of text that should cause the column to expand automatically");
                worksheet.Cells["A2"].PutValue(12345);
                worksheet.Cells["B2"].PutValue(DateTime.Now);
                worksheet.Cells["C2"].PutValue("Another long text entry to test column width adjustment");

                // Auto‑fit all columns based on the content in the worksheet
                worksheet.AutoFitColumns();

                // Display the new column widths in the console
                Console.WriteLine("Column A width after AutoFit: " + worksheet.Cells.GetColumnWidth(0));
                Console.WriteLine("Column B width after AutoFit: " + worksheet.Cells.GetColumnWidth(1));
                Console.WriteLine("Column C width after AutoFit: " + worksheet.Cells.GetColumnWidth(2));

                // Define output file path
                string outputPath = "AutoFitColumnsResult.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (lifecycle: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during workbook processing: {ex.Message}");
            }
        }
    }
}