// Title: Parallel batch conversion of Excel workbooks to PDF with Aspose.Cells in C#
// Common Searches: C# how to convert a list of .xlsx files to PDF using Aspose.Cells with Parallel.ForEach | example of multithreaded Excel to PDF conversion using Aspose.Cells .NET | batch process Excel workbooks to PDF in parallel for high throughput | Aspose.Cells parallel conversion of multiple Excel files to PDF in .NET Core | speed up Excel to PDF conversion using C# Parallel.ForEach and Aspose.Cells
// Tags: Aspose.Cells parallel Excel to PDF conversion | C# batch workbook PDF generation Aspose.Cells | multithreaded Excel file conversion .NET | Parallel.ForEach Excel to PDF processing | save workbook as PDF using Aspose.Cells C# | high‑throughput Excel PDF conversion .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace ExcelToPdfBatch
{
    // The sample lists Excel file paths, creates an output folder, and uses Parallel.ForEach with a degree of parallelism equal to the processor count to load each workbook via Aspose.Cells and save it as a PDF. It logs successes and errors for each file and reports when the batch conversion is complete.
    class Program
    {
        static void Main(string[] args)
        {
            // Example: list of Excel files to convert.
            // In a real scenario, you might read these from a directory or an input source.
            List<string> excelFiles = new List<string>
            {
                @"C:\Input\Report1.xlsx",
                @"C:\Input\Report2.xls",
                @"C:\Input\Report3.xlsx"
                // Add more file paths as needed.
            };

            // Destination folder for PDF files.
            string outputFolder = @"C:\Output\PDFs";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Use Parallel.ForEach to process files concurrently.
            ParallelOptions parallelOptions = new ParallelOptions
            {
                // Adjust the degree of parallelism based on the environment.
                // Setting it to Environment.ProcessorCount is a common choice.
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };

            Parallel.ForEach(excelFiles, parallelOptions, excelPath =>
            {
                try
                {
                    // Load the Excel workbook.
                    Workbook workbook = new Workbook(excelPath);

                    // Determine the output PDF file name.
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                    string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Save the workbook as PDF.
                    workbook.Save(pdfPath, SaveFormat.Pdf);

                    Console.WriteLine($"Successfully converted: {excelPath} -> {pdfPath}");
                }
                catch (Exception ex)
                {
                    // Log the error; in production you might write to a log file.
                    Console.Error.WriteLine($"Error converting '{excelPath}': {ex.Message}");
                }
            });

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
