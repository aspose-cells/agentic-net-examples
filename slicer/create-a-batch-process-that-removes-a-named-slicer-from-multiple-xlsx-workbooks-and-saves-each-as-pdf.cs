// Title: Batch remove a named slicer from multiple XLSX workbooks and save each as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# console application that scans a folder for .xlsx files, deletes a slicer with a given name from every worksheet using the Aspose.Cells.Slicers API, and writes the result as a PDF to an output directory. | Generate C# code that loads each Excel workbook in a directory, removes the slicer called "MySlicer" from all sheets via Aspose.Cells, and exports the cleaned workbook to a PDF file with the same base name.
// Common Searches: aspnet remove slicer from excel files in bulk using Aspose.Cells | c# batch process to delete a specific slicer and convert workbooks to pdf | how to programmatically remove a named slicer from all worksheets with Aspose.Cells | convert multiple xlsx to pdf after removing slicer using Aspose.Cells C# | automate slicer cleanup and pdf export for a folder of Excel workbooks
// Tags: Aspose.Cells slicer removal C# | batch convert xlsx to pdf Aspose.Cells | delete specific slicer from worksheets | automate Excel workbook PDF export | process multiple workbooks with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Correct namespace for Slicer

// The C# utility iterates through every .xlsx file in a specified input folder, loads each workbook with Aspose.Cells, removes a slicer named "MySlicer" from all worksheets, and saves the modified workbook as a PDF in an output folder, handling missing directories and errors gracefully.
class BatchSlicerRemoval
{
    static void Main()
    {
        // Folder containing the source XLSX workbooks
        string inputFolder = @"C:\Input";

        // Folder where the resulting PDF files will be saved
        string outputFolder = @"C:\Output";

        // Name of the slicer to be removed from each workbook
        string slicerName = "MySlicer";

        // Ensure the output directory exists
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
        }

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Process each XLSX file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Guard against missing files (should not happen in GetFiles loop)
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Iterate through all worksheets to locate and remove the slicer
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Attempt to retrieve the slicer by its name
                    Slicer slicer = sheet.Slicers[slicerName];

                    // If the slicer exists, remove it from the worksheet
                    if (slicer != null)
                    {
                        sheet.Slicers.Remove(slicer);
                    }
                }

                // Construct the PDF file name (same base name as the source workbook)
                string pdfPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                // Save the modified workbook as a PDF document
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"Processed and saved PDF: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
