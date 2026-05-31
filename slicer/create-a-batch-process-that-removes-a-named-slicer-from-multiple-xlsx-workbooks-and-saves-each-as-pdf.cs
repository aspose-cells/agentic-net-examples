using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Utility;

namespace BatchSlicerRemovalApp
{
    class BatchSlicerRemoval
    {
        static void Main()
        {
            // Folder containing the source XLSX workbooks
            string sourceFolder = @"C:\Input";
            // Folder where the resulting PDFs will be saved
            string outputFolder = @"C:\Output";
            // Name of the slicer to be removed from each workbook
            string slicerNameToRemove = "FruitSlicer1";

            try
            {
                // Verify source folder exists
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder not found: {sourceFolder}");
                    return;
                }

                // Ensure output folder exists
                Directory.CreateDirectory(outputFolder);

                // Process each XLSX file in the source folder
                foreach (string xlsxPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
                {
                    // Verify the file exists before loading
                    if (!File.Exists(xlsxPath))
                    {
                        Console.WriteLine($"File not found (skipped): {xlsxPath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(xlsxPath);

                        // Iterate through all worksheets
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            // Get the slicer collection for the current worksheet
                            SlicerCollection slicers = sheet.Slicers;

                            // Iterate backwards for safe removal
                            for (int i = slicers.Count - 1; i >= 0; i--)
                            {
                                Slicer slicer = slicers[i];
                                if (slicer.Name == slicerNameToRemove)
                                {
                                    slicers.Remove(slicer);
                                }
                            }
                        }

                        // Save the modified workbook to a temporary file
                        string tempXlsxPath = Path.Combine(outputFolder,
                            Path.GetFileNameWithoutExtension(xlsxPath) + "_temp.xlsx");
                        workbook.Save(tempXlsxPath, SaveFormat.Xlsx);

                        // Convert the temporary XLSX file to PDF
                        string pdfPath = Path.Combine(outputFolder,
                            Path.GetFileNameWithoutExtension(xlsxPath) + ".pdf");
                        ConversionUtility.Convert(tempXlsxPath, pdfPath);

                        // Delete the temporary file
                        if (File.Exists(tempXlsxPath))
                        {
                            File.Delete(tempXlsxPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{xlsxPath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}