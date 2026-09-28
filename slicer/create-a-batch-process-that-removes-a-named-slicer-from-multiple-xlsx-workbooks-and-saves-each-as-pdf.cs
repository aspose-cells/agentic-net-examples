// Title: Batch remove a specific slicer from multiple XLSX workbooks and save each as PDF with Aspose.Cells for .NET
// AI Prompts: Write a C# console program that loads a list of .xlsx files, deletes the slicer named 'FruitSlicer1' from every worksheet, and converts the modified workbook to PDF using Aspose.Cells ConversionUtility. | Generate C# code that iterates through a collection of Excel workbooks, removes a targeted slicer on all sheets, saves each workbook to a temporary file, and exports it to PDF in a designated output folder. | Create a C# script that validates each input workbook, removes a named slicer, handles missing files gracefully, and cleans up temporary files after PDF conversion with Aspose.Cells.
// Common Searches: asp.net remove slicer from multiple excel files programmatically | batch delete specific slicer in Aspose.Cells and export to PDF | C# code to loop through XLSX workbooks, remove slicer and convert to PDF | how to use ConversionUtility to convert modified workbook to PDF after slicer removal | automate slicer removal and PDF generation for many Excel reports in .NET
// Tags: Aspose.Cells remove slicer batch processing | C# delete slicer from worksheets | XLSX to PDF conversion with ConversionUtility | temporary workbook handling Aspose.Cells | bulk Excel slicer removal .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Utility;

namespace BatchSlicerRemoval
{
    // // This C# console application iterates over a predefined array of .xlsx files, removes the slicer named 'FruitSlicer1' from every worksheet using the Aspose.Cells SlicerCollection, saves the edited workbook to a temporary file, converts it to PDF with ConversionUtility, writes the PDF to an output folder, and deletes the temporary file.
    class Program
    {
        static void Main()
        {
            // Define the slicer name to remove
            string slicerNameToRemove = "FruitSlicer1";

            // Input workbook files (XLSX) – adjust the paths as needed
            string[] workbookFiles = new string[]
            {
                @"C:\Input\Report1.xlsx",
                @"C:\Input\Report2.xlsx",
                @"C:\Input\Report3.xlsx"
            };

            // Output folder for the generated PDFs
            string outputFolder = @"C:\Output";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each workbook
            foreach (string workbookPath in workbookFiles)
            {
                // Verify the input file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"File not found: {workbookPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook wb = new Workbook(workbookPath);

                    // Iterate through all worksheets
                    foreach (Worksheet ws in wb.Worksheets)
                    {
                        // Get the slicer collection for the current worksheet
                        SlicerCollection slicers = ws.Slicers;

                        // Find the slicer with the specified name
                        Slicer targetSlicer = null;
                        foreach (Slicer s in slicers)
                        {
                            if (s.Name.Equals(slicerNameToRemove, StringComparison.OrdinalIgnoreCase))
                            {
                                targetSlicer = s;
                                break;
                            }
                        }

                        // If found, remove it using the collection's Remove method
                        if (targetSlicer != null)
                        {
                            slicers.Remove(targetSlicer);
                        }
                    }

                    // Save the modified workbook to a temporary XLSX file
                    string tempXlsxPath = Path.Combine(Path.GetTempPath(),
                        Guid.NewGuid().ToString() + ".xlsx");
                    wb.Save(tempXlsxPath, SaveFormat.Xlsx);

                    // Determine the PDF output path (same name as original workbook)
                    string pdfFileName = Path.GetFileNameWithoutExtension(workbookPath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Convert the temporary XLSX to PDF using ConversionUtility
                    ConversionUtility.Convert(tempXlsxPath, pdfPath);

                    // Clean up the temporary XLSX file
                    if (File.Exists(tempXlsxPath))
                    {
                        File.Delete(tempXlsxPath);
                    }

                    Console.WriteLine($"Processed '{workbookPath}' → '{pdfPath}'");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
                }
            }
        }
    }
}
