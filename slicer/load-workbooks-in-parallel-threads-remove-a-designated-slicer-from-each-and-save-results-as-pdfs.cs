// Title: Remove a named slicer from each worksheet of multiple Excel files in parallel and save them as PDFs using Aspose.Cells for .NET
// AI Prompts: Write C# code that enumerates every .xlsx file in a folder, loads each workbook with Aspose.Cells, deletes the slicer whose Name equals "MySlicer" from all worksheets, and saves the result as a PDF in a separate output directory using Parallel.ForEach for concurrency. | Generate a multithreaded Aspose.Cells solution that scans a source directory, removes a specific slicer from every worksheet of each workbook, and converts the modified workbooks to PDF files, handling missing files and ensuring the output folder exists.
// Common Searches: asp.net remove slicer from all worksheets using Aspose.Cells | parallel processing multiple Excel files to PDF with Aspose.Cells C# | batch delete specific slicer and export workbooks to PDF Aspose.Cells .NET | how to use Parallel.ForEach with Aspose.Cells to modify and save workbooks
// Tags: Aspose.Cells remove slicer C# | parallel workbook processing .NET | batch Excel to PDF conversion Aspose | slicer collection manipulation Aspose.Cells | multithreaded PDF export Aspose.Cells

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for Slicer and SlicerCollection

// The program scans a directory for .xlsx files, loads each workbook with Aspose.Cells, removes the slicer named "MySlicer" from every worksheet, and saves the modified workbook as a PDF in an output folder, processing all files concurrently with Parallel.ForEach.
class Program
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where the resulting PDFs will be saved
        string outputFolder = @"C:\OutputPdfs";
        // Name of the slicer to be removed from each workbook
        string slicerNameToRemove = "MySlicer";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Get all Excel files in the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        // Process each workbook in parallel
        Parallel.ForEach(excelFiles, excelFilePath =>
        {
            try
            {
                // Guard against missing files (unlikely in GetFiles, but safe for robustness)
                if (!File.Exists(excelFilePath))
                {
                    Console.WriteLine($"File not found: {excelFilePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(excelFilePath);

                // Iterate through all worksheets to find and remove the designated slicer
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    SlicerCollection slicers = sheet.Slicers;
                    // Iterate backwards when removing items from a collection
                    for (int i = slicers.Count - 1; i >= 0; i--)
                    {
                        Slicer slicer = slicers[i];
                        if (slicer.Name.Equals(slicerNameToRemove, StringComparison.OrdinalIgnoreCase))
                        {
                            slicers.RemoveAt(i);
                        }
                    }
                }

                // Build the output PDF file path
                string pdfFileName = Path.GetFileNameWithoutExtension(excelFilePath) + ".pdf";
                string pdfFilePath = Path.Combine(outputFolder, pdfFileName);

                // Save the modified workbook as a PDF
                workbook.Save(pdfFilePath, SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{excelFilePath}': {ex.Message}");
            }
        });

        Console.WriteLine("Processing completed.");
    }
}
