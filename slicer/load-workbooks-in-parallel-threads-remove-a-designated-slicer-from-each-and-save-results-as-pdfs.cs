// Title: Remove a named slicer from multiple Excel workbooks concurrently and export each workbook to PDF using Aspose.Cells for .NET
// Common Searches: c# aspnet remove slicer from excel workbook using Aspose.Cells | parallel processing of multiple Excel files with Aspose.Cells and save as PDF | how to delete a specific slicer in all worksheets with Aspose.Cells C# | batch convert Excel to PDF after removing slicers using Aspose.Cells | using Parallel.ForEach to modify Excel workbooks with Aspose.Cells | Aspose.Cells remove slicer by name and export to PDF in .NET
// Tags: remove slicer Aspose.Cells C# | parallel workbook processing Aspose.Cells | export workbook to PDF Aspose.Cells | batch slicer deletion Excel Aspose.Cells | multithreaded Excel to PDF conversion Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for Slicer class

namespace AsposeCellsParallelSlicerRemoval
{
    // The example loads a list of Excel files, processes each workbook in parallel with Parallel.ForEach, removes any slicer named "MySlicer" from all worksheets, and saves the modified workbook as a PDF in the same folder using Aspose.Cells for .NET.
    class Program
    {
        // Designated slicer name to be removed from each workbook
        private const string SlicerNameToRemove = "MySlicer";

        static void Main(string[] args)
        {
            // Example: list of Excel files to process
            List<string> excelFiles = new List<string>
            {
                @"C:\Input\Book1.xlsx",
                @"C:\Input\Book2.xlsx",
                @"C:\Input\Book3.xlsx"
                // Add more file paths as needed
            };

            // Process each workbook in parallel
            Parallel.ForEach(excelFiles, filePath =>
            {
                try
                {
                    // Verify that the input file exists
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: '{filePath}'. Skipping.");
                        return;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Iterate through all worksheets to find and remove the slicer
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Collect slicers that match the designated name
                        List<Slicer> slicersToRemove = new List<Slicer>();
                        foreach (Slicer slicer in sheet.Slicers)
                        {
                            if (string.Equals(slicer.Name, SlicerNameToRemove, StringComparison.OrdinalIgnoreCase))
                            {
                                slicersToRemove.Add(slicer);
                            }
                        }

                        // Remove the identified slicers
                        foreach (Slicer slicer in slicersToRemove)
                        {
                            sheet.Slicers.Remove(slicer);
                        }
                    }

                    // Prepare PDF output path (same folder, same name with .pdf extension)
                    string pdfPath = Path.ChangeExtension(filePath, ".pdf");

                    // Ensure the output directory exists
                    string outputDir = Path.GetDirectoryName(pdfPath);
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the modified workbook as PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);
                    Console.WriteLine($"Processed and saved PDF: '{pdfPath}'.");
                }
                catch (Exception ex)
                {
                    // Log or handle errors per file as needed
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            });

            Console.WriteLine("Processing completed.");
        }
    }
}
