using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells;

namespace OleObjectExtractionDemo
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook;
            try
            {
                workbook = new Workbook("InputWithOleObjects.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all OLE objects in the current worksheet
                foreach (OleObject ole in sheet.OleObjects)
                {
                    // Skip linked OLE objects – they do not contain embedded data
                    if (ole.IsLink)
                    {
                        Console.WriteLine("Skipping linked OLE object.");
                        continue;
                    }

                    try
                    {
                        // Attempt to retrieve the embedded OLE data.
                        // Using the ObjectData property (rule: OleObject.ObjectData)
                        byte[] data = ole.ObjectData;

                        // If ObjectData is null or empty, try FullObjectBin as a fallback
                        if (data == null || data.Length == 0)
                        {
                            // FullObjectBin is read‑only (rule: OleObject.FullObjectBin)
                            data = ole.FullObjectBin;
                        }

                        // If we still have no data, report and continue
                        if (data == null || data.Length == 0)
                        {
                            Console.WriteLine("OLE object contains no data.");
                            continue;
                        }

                        // Save the extracted OLE data to a file for verification
                        string outputFileName = $"OleObject_{sheet.Name}_{ole.Name}_{Guid.NewGuid()}.bin";
                        File.WriteAllBytes(outputFileName, data);
                        Console.WriteLine($"Extracted OLE object saved to: {outputFileName}");
                    }
                    catch (CellsException cex) when (cex.Code == ExceptionType.FileCorrupted ||
                                                    cex.Code == ExceptionType.UnsupportedFeature ||
                                                    cex.Code == ExceptionType.UnsupportedStream)
                    {
                        // Specific handling for corrupted or unsupported OLE objects
                        Console.WriteLine($"Unable to extract OLE object (corrupted/unsupported). Details: {cex.Message}");
                    }
                    catch (Exception ex)
                    {
                        // General fallback for any other unexpected errors
                        Console.WriteLine($"Unexpected error while processing OLE object: {ex.Message}");
                    }
                }
            }

            // Optionally, save the workbook after processing (if any modifications were made)
            try
            {
                workbook.Save("ProcessedWorkbook.xlsx");
                Console.WriteLine("Workbook saved as ProcessedWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}