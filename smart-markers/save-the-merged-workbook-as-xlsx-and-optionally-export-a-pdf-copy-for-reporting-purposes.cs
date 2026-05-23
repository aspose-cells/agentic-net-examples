using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorkbookSaver
    {
        /// <summary>
        /// Saves the provided workbook as XLSX and optionally as PDF.
        /// </summary>
        /// <param name="workbook">The workbook to be saved.</param>
        /// <param name="exportPdf">If true, also saves a PDF copy.</param>
        public static void SaveWorkbook(Workbook workbook, bool exportPdf = false)
        {
            try
            {
                // Save the workbook in XLSX format
                workbook.Save("MergedWorkbook.xlsx", SaveFormat.Xlsx);

                // Optionally export a PDF version
                if (exportPdf)
                {
                    workbook.Save("MergedWorkbook.pdf", SaveFormat.Pdf);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }

        // Example usage
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook mergedWorkbook = new Workbook();

                // Populate with sample data (replace with actual merge logic)
                Worksheet ws = mergedWorkbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Merged Data");

                // Save as XLSX and also generate a PDF report
                SaveWorkbook(mergedWorkbook, exportPdf: true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during Run: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookSaver.Run();
        }
    }
}