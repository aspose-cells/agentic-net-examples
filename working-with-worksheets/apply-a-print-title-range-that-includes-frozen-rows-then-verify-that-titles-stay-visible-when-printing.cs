using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class PrintTitleWithFrozenRowsDemo
    {
        public static void Run()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // ---------- Populate sample data ----------
                // Row 1 will be the header (title row)
                worksheet.Cells["A1"].PutValue("Header");

                // Add many data rows to force pagination
                for (int i = 2; i <= 200; i++)
                {
                    worksheet.Cells[$"A{i}"].PutValue($"Data {i - 1}");
                }

                // ---------- Freeze the first row ----------
                // Freeze at cell A2 (row index 1) with 1 frozen row and 0 frozen columns
                worksheet.FreezePanes(1, 0, 1, 0);

                // ---------- Set print title rows ----------
                // The frozen header row should repeat on each printed page
                worksheet.PageSetup.PrintTitleRows = "$1:$1";

                // Optional: define a print area covering all data
                worksheet.PageSetup.PrintArea = "A1:A200";

                // ---------- Verify that the title row is included in page breaks ----------
                // Create print options (default)
                ImageOrPrintOptions printOptions = new ImageOrPrintOptions();

                // Retrieve automatic page breaks based on the current page setup
                CellArea[] pageBreaks = worksheet.GetPrintingPageBreaks(printOptions);

                // The first page break should start at row 0 (which is the title row)
                bool titleInFirstPage = false;
                if (pageBreaks != null && pageBreaks.Length > 0)
                {
                    // The first page area starts at row 0 and ends at pageBreaks[0].EndRow
                    titleInFirstPage = pageBreaks[0].StartRow <= 0 && pageBreaks[0].EndRow >= 0;
                }

                Console.WriteLine($"PrintTitleRows set to: {worksheet.PageSetup.PrintTitleRows}");
                Console.WriteLine($"FreezePanes applied: Row 1 frozen");
                Console.WriteLine($"Title row appears in first printed page: {titleInFirstPage}");

                // ---------- Save the workbook ----------
                string outputPath = "PrintTitleWithFrozenRowsDemo.xlsx";

                // Ensure we don't attempt to overwrite a read‑only file
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            PrintTitleWithFrozenRowsDemo.Run();
        }
    }
}