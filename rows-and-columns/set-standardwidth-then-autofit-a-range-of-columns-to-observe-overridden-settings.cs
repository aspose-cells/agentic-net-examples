using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class StandardWidthAutoFitDemo
    {
        public static void Main()
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
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set the default column width (StandardWidth) to 20 characters
            cells.StandardWidth = 20.0;
            Console.WriteLine($"StandardWidth set to: {cells.StandardWidth}");

            // Manually set column widths for columns A, B, C
            cells.SetColumnWidth(0, 10.0); // Column A narrower than standard
            cells.SetColumnWidth(1, 30.0); // Column B wider than standard
            cells.SetColumnWidth(2, 15.0); // Column C different

            // Display initial widths
            Console.WriteLine($"Initial Column A width: {cells.GetColumnWidth(0)}");
            Console.WriteLine($"Initial Column B width: {cells.GetColumnWidth(1)}");
            Console.WriteLine($"Initial Column C width: {cells.GetColumnWidth(2)}");

            // Populate cells with long text to require auto‑fit
            worksheet.Cells["A1"].PutValue("This is a very long text that should cause column A to expand when auto‑fit is applied.");
            worksheet.Cells["B1"].PutValue("Short");
            worksheet.Cells["C1"].PutValue("Another extremely long piece of text that will trigger auto‑fit for column C.");

            // Auto‑fit columns A to C (indexes 0 to 2)
            worksheet.AutoFitColumns(0, 2);

            // Display widths after auto‑fit
            Console.WriteLine($"After AutoFit - Column A width: {cells.GetColumnWidth(0)}");
            Console.WriteLine($"After AutoFit - Column B width: {cells.GetColumnWidth(1)}");
            Console.WriteLine($"After AutoFit - Column C width: {cells.GetColumnWidth(2)}");

            // Save the workbook
            string outputPath = "StandardWidthAutoFitDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}