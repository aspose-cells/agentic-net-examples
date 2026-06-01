using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HideRowsAndSavePdf
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the existing Excel file
            string sourcePath = "input.xlsx";

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                throw new FileNotFoundException($"Source file not found: {sourcePath}");
            }

            // Load the workbook from the source Excel file
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide rows 20 to 25 (zero‑based index: start at 19, hide 6 rows)
            worksheet.Cells.HideRows(19, 6);

            // Save the modified workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);
        }
    }
}