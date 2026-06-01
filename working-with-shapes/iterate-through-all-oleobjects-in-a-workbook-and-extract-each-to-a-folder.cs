using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectExtractor
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source workbook
                string sourceFile = @"C:\Input\WorkbookWithOleObjects.xlsx";

                // Verify that the workbook file exists
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file not found: {sourceFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourceFile);

                // Destination folder for extracted OLE objects
                string outputFolder = @"C:\Output\OleObjects";
                Directory.CreateDirectory(outputFolder);

                int extractedCount = 0;

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each OLE object in the worksheet
                    foreach (OleObject ole in sheet.OleObjects)
                    {
                        try
                        {
                            string filePath;

                            if (ole.IsLink && !string.IsNullOrEmpty(ole.ObjectSourceFullName))
                            {
                                // Linked OLE object – copy the source file if it exists
                                string sourcePath = ole.ObjectSourceFullName;
                                if (!File.Exists(sourcePath))
                                {
                                    Console.WriteLine($"Linked source not found: {sourcePath}");
                                    continue;
                                }

                                string fileName = Path.GetFileName(sourcePath);
                                filePath = Path.Combine(outputFolder, $"Linked_{extractedCount}_{fileName}");
                                File.Copy(sourcePath, filePath, overwrite: true);
                            }
                            else
                            {
                                // Embedded OLE object – write the binary data to a file
                                byte[] data = ole.ObjectData;
                                if (data == null || data.Length == 0)
                                {
                                    // No data to extract; skip this object
                                    continue;
                                }

                                // Determine file extension based on the format type (fallback to .bin)
                                string extension = ".bin";
                                try
                                {
                                    extension = "." + ole.FileFormatType.ToString().ToLower();
                                }
                                catch { /* ignore and keep default */ }

                                filePath = Path.Combine(outputFolder, $"Embedded_{extractedCount}{extension}");
                                File.WriteAllBytes(filePath, data);
                            }

                            Console.WriteLine($"Extracted OLE object to: {filePath}");
                            extractedCount++;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error extracting OLE object: {ex.Message}");
                        }
                    }
                }

                // Optionally, save the workbook if any modifications were made
                // workbook.Save(@"C:\Output\ModifiedWorkbook.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}