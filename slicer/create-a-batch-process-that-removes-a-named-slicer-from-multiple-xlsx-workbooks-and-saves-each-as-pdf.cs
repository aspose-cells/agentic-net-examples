// Title: C# batch removal of a named slicer from multiple Excel workbooks and PDF export with Aspose.Cells
// Description: A .NET console app that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, deletes every slicer whose Name matches a specified value, saves the cleaned workbook temporarily, converts it to PDF, and cleans up the intermediate file. Ideal for bulk preparation of Excel reports before distribution.
// Keywords: Aspose.Cells slicer removal C# | batch delete slicer Excel | convert Excel to PDF Aspose.Cells | C# process multiple workbooks | remove slicer programmatically | Excel slicer automation | Aspose.Cells PDF conversion | bulk Excel cleanup
// Common Searches: How to delete a specific slicer from many Excel files using Aspose.Cells | Batch convert Excel workbooks to PDF after removing slicers in C# | Remove slicer named MySlicer from all worksheets programmatically | Aspose.Cells example for bulk slicer removal and PDF export | C# script to strip slicers from Excel and generate PDFs
// Developer Intent: Automatically strip a given slicer from every worksheet in each workbook within a directory and produce a PDF version of the cleaned file.
// Use Cases: Prepare reporting packages by removing interactive slicers before sending PDFs to clients. | Archive Excel dashboards without slicer controls for compliance or storage efficiency. | Process incoming workbooks to enforce policy‑based slicer removal and generate review‑ready PDFs.
// AI Prompts: Generate a reusable method that accepts source folder, slicer name, and output folder, removes matching slicers from all worksheets, and saves each workbook as PDF using Aspose.Cells. | Add comprehensive error handling and logging to the batch slicer removal script, capturing load failures, missing slicers, and PDF conversion issues. | Refactor the code to convert workbooks directly to PDF without creating a temporary XLSX file, preserving performance.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Utility;

namespace BatchSlicerRemoval
{
    // A .NET console app that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, deletes every slicer whose Name matches a specified value, saves the cleaned workbook temporarily, converts it to PDF, and cleans up the intermediate file. Ideal for bulk preparation of Excel reports before distribution.
    class Program
    {
        static void Main()
        {
            // Folder containing the source XLSX workbooks
            string sourceFolder = @"C:\Input";
            // Folder where the resulting PDFs will be saved
            string outputFolder = @"C:\Output";

            // Name of the slicer to be removed from each workbook
            string slicerNameToRemove = "MySlicer";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each XLSX file in the source folder
            foreach (string xlsxPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                // Load the workbook (lifecycle: load)
                Workbook workbook = new Workbook(xlsxPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the slicer collection for the current worksheet
                    SlicerCollection slicers = sheet.Slicers;

                    // Iterate backwards to safely remove items while iterating
                    for (int i = slicers.Count - 1; i >= 0; i--)
                    {
                        Slicer slicer = slicers[i];
                        // Remove the slicer if its name matches the target name
                        if (slicer.Name == slicerNameToRemove)
                        {
                            slicers.Remove(slicer); // Use SlicerCollection.Remove method
                        }
                    }
                }

                // Save the modified workbook as a temporary XLSX file (lifecycle: save)
                string tempXlsxPath = Path.Combine(
                    outputFolder,
                    Path.GetFileNameWithoutExtension(xlsxPath) + "_temp.xlsx");
                workbook.Save(tempXlsxPath, SaveFormat.Xlsx);

                // Convert the temporary XLSX to PDF using the provided ConversionUtility (rule)
                string pdfPath = Path.Combine(
                    outputFolder,
                    Path.GetFileNameWithoutExtension(xlsxPath) + ".pdf");
                ConversionUtility.Convert(tempXlsxPath, pdfPath);

                // Optionally delete the temporary XLSX file
                File.Delete(tempXlsxPath);
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
