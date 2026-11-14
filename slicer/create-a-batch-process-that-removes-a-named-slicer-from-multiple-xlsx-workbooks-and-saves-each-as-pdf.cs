// Title: Batch removal of a named slicer from multiple Excel workbooks and PDF conversion using Aspose.Cells for .NET
// Common Searches: C# code to delete a specific slicer from all worksheets in a folder of .xlsx files using Aspose.Cells | Aspose.Cells batch processing to remove slicer and export workbooks to PDF | Remove slicer by name from multiple Excel files programmatically with Aspose.Cells .NET | Convert modified Excel workbooks to PDF after deleting slicers using Aspose.Cells | Iterate through worksheets and delete slicer collection items in Aspose.Cells C# | Automate slicer removal and PDF conversion for many Excel files in .NET
// Tags: Aspose.Cells slicer removal batch | C# delete slicer from worksheet collection | export modified workbook to PDF Aspose.Cells | process multiple XLSX files programmatically | slicer collection iteration Aspose.Cells | automated Excel to PDF conversion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for SlicerCollection and Slicer classes

namespace SlicerRemovalBatch
{
    // The example scans an input folder for .xlsx files, loads each workbook with Aspose.Cells, iterates through all worksheets to locate and delete a slicer matching a given name, then saves the updated workbook as a PDF in an output folder, handling missing files and exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source XLSX workbooks
            string inputFolder = @"C:\Input";

            // Folder where the resulting PDF files will be saved
            string outputFolder = @"C:\Output";

            // Name of the slicer to be removed from each workbook
            string slicerName = "MySlicer";

            // Ensure the input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Process each XLSX file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Verify the file still exists before attempting to load
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Iterate through all worksheets to locate and remove the named slicer
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        SlicerCollection slicers = sheet.Slicers;

                        // Iterate backwards when removing items from a collection
                        for (int i = slicers.Count - 1; i >= 0; i--)
                        {
                            if (slicers[i].Name.Equals(slicerName, StringComparison.OrdinalIgnoreCase))
                            {
                                slicers.RemoveAt(i);
                                // If only one slicer with this name exists, you can break here
                            }
                        }
                    }

                    // Build the PDF file name (same base name as the source workbook)
                    string pdfFilePath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                    // Save the modified workbook as PDF
                    workbook.Save(pdfFilePath, SaveFormat.Pdf);
                    Console.WriteLine($"Processed and saved PDF: {pdfFilePath}");
                }
                catch (Exception ex)
                {
                    // Log the error and continue with the next file
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
