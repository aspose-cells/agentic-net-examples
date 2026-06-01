using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class WorkbookToPdfConverter
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new (empty) workbook
                Workbook workbook = new Workbook();

                // Set each worksheet to left‑to‑right display direction
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.DisplayRightToLeft = false;
                }

                string outputPath = "output.pdf";

                // Save the workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"Workbook successfully converted to PDF at '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}