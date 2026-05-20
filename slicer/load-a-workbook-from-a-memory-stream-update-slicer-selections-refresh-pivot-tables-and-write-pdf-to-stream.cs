using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsDemo
{
    public class Program
    {
        static void Main(string[] args)
        {
            // Paths for input Excel file and output PDF file
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the Excel file into a memory stream
                using (FileStream fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                using (MemoryStream excelStream = new MemoryStream())
                {
                    fileStream.CopyTo(excelStream);

                    // Convert to PDF with slicer and pivot table refresh
                    MemoryStream pdfStream = WorkbookProcessor.ConvertToPdfWithRefresh(excelStream);
                    if (pdfStream == null)
                    {
                        Console.WriteLine("Conversion failed.");
                        return;
                    }

                    // Save the resulting PDF to disk
                    using (FileStream outStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        pdfStream.CopyTo(outStream);
                    }

                    Console.WriteLine($"PDF successfully saved to: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class WorkbookProcessor
    {
        /// <summary>
        /// Loads an Excel workbook from the provided memory stream, refreshes all slicers and pivot tables,
        /// and returns the workbook saved as a PDF in a new memory stream.
        /// </summary>
        /// <param name="excelStream">MemoryStream containing the source Excel file.</param>
        /// <returns>MemoryStream containing the PDF representation of the workbook, or null if an error occurs.</returns>
        public static MemoryStream ConvertToPdfWithRefresh(MemoryStream excelStream)
        {
            try
            {
                // Ensure the input stream is positioned at the beginning
                excelStream.Position = 0;

                // Load the workbook from the stream
                Workbook workbook = new Workbook(excelStream);

                // Refresh all slicers in all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    for (int i = 0; i < sheet.Slicers.Count; i++)
                    {
                        Slicer slicer = sheet.Slicers[i];
                        slicer.Refresh(); // Refresh slicer and its underlying pivot tables
                    }
                }

                // Refresh all pivot tables in the workbook
                workbook.Worksheets.RefreshPivotTables();

                // Save the refreshed workbook as PDF into a new memory stream
                MemoryStream pdfStream = new MemoryStream();
                workbook.Save(pdfStream, SaveFormat.Pdf);
                pdfStream.Position = 0; // Reset position for reading

                return pdfStream;
            }
            catch (Exception ex)
            {
                // Log the exception (could be replaced with proper logging)
                Console.WriteLine($"Conversion error: {ex.Message}");
                return null;
            }
        }
    }
}