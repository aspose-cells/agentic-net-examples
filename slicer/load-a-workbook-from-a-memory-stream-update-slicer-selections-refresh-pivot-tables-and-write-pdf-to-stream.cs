using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsPdfDemo
{
    public static class ExcelToPdfConverter
    {
        /// <summary>
        /// Loads an Excel workbook from a memory stream, refreshes all slicers and pivot tables,
        /// and returns a PDF representation of the workbook in a new memory stream.
        /// </summary>
        /// <param name="excelStream">MemoryStream containing the source Excel file.</param>
        /// <returns>MemoryStream containing the PDF output.</returns>
        public static MemoryStream ConvertExcelToPdfWithSlicerRefresh(MemoryStream excelStream)
        {
            try
            {
                // Ensure the input stream is positioned at the beginning
                excelStream.Position = 0;

                // Load the workbook from the provided stream
                using (Workbook workbook = new Workbook(excelStream))
                {
                    // Refresh all slicers in every worksheet (if any)
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        for (int i = 0; i < sheet.Slicers.Count; i++)
                        {
                            Slicer slicer = sheet.Slicers[i];
                            slicer.Refresh(); // Refresh slicer and its underlying pivot tables
                        }
                    }

                    // Refresh all pivot tables in the workbook (covers any that might not be linked to slicers)
                    workbook.Worksheets.RefreshPivotTables();

                    // Save the refreshed workbook as PDF into a new memory stream
                    MemoryStream pdfStream = new MemoryStream();
                    workbook.Save(pdfStream, SaveFormat.Pdf);
                    pdfStream.Position = 0; // Reset for downstream reading

                    return pdfStream;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
                throw; // Re‑throw to let the caller handle if needed
            }
        }
    }

    public class Program
    {
        /// <summary>
        /// Entry point: converts an Excel file to PDF while refreshing slicers.
        /// Usage: AsposeCellsPdfDemo.exe <input.xlsx> <output.pdf>
        /// </summary>
        static void Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: <input.xlsx> <output.pdf>");
                    return;
                }

                string inputPath = args[0];
                string outputPath = args[1];

                // Prevent FileNotFoundException for the input file
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the Excel file into a memory stream
                using (FileStream fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                using (MemoryStream excelStream = new MemoryStream())
                {
                    fileStream.CopyTo(excelStream);

                    // Convert to PDF
                    MemoryStream pdfStream = ExcelToPdfConverter.ConvertExcelToPdfWithSlicerRefresh(excelStream);

                    // Write PDF to the output file
                    using (FileStream outStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        pdfStream.CopyTo(outStream);
                    }

                    Console.WriteLine($"PDF successfully saved to: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error: {ex.Message}");
            }
        }
    }
}