// Title: C# Example: Extract OLE Objects from Excel with Aspose.Cells – Robust Error Handling for Corrupted and Unsupported Formats
// Description: Loads an Excel file, iterates all worksheets and their OleObjects, extracts each object's FullObjectBin to a separate .bin file, and catches CellsException for FileCorrupted and UnsupportedFeature while logging unexpected errors. The workbook is saved after processing, ensuring the program continues even when some OLE objects are damaged or unsupported.
// Keywords: Aspose.Cells OLE extraction C# | extract OLE objects Excel | handle corrupted OLE Aspose | unsupported OLE format exception | FullObjectBin error handling | C# Excel OLE binary export | Aspose.Cells try catch example | Excel embedded file extraction | robust OLE processing .NET | GitHub Aspose.Cells OLE sample
// Common Searches: how to extract OLE objects from Excel using Aspose.Cells C# | catch CellsException for corrupted OLE objects | unsupported OLE format handling Aspose.Cells | save OLE binary data to file C# | Aspose.Cells example for OLE extraction with error handling
// Developer Intent: Implement try‑catch blocks to safely extract OLE objects from a workbook, handle corrupted or unsupported OLE data, and continue processing without interruption.
// Use Cases: Batch‑process spreadsheets to pull embedded Word, PDF, or image files while skipping damaged OLE objects. | Create an audit log of all extracted OLE binaries for compliance reporting, with graceful handling of unsupported formats. | Integrate OLE extraction into a data‑migration pipeline that must not fail when encountering corrupted embedded objects.
// AI Prompts: Write C# code that uses Aspose.Cells to extract OLE objects from an Excel file and logs FileCorrupted and UnsupportedFeature exceptions to a text file. | Refactor the provided program to use a configurable logger (e.g., NLog) instead of console output while preserving error‑handling logic. | Suggest a retry strategy for transient I/O errors when reading FullObjectBin data from OLE objects in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectExtractionWithErrorHandling
{
    // Loads an Excel file, iterates all worksheets and their OleObjects, extracts each object's FullObjectBin to a separate .bin file, and catches CellsException for FileCorrupted and UnsupportedFeature while logging unexpected errors. The workbook is saved after processing, ensuring the program continues even when some OLE objects are damaged or unsupported.
    class Program
    {
        static void Main()
        {
            const string inputPath = "InputWithOleObjects.xlsx";
            const string outputWorkbookPath = "ProcessedWorkbook.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            int oleCounter = 0;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all OLE objects in the current worksheet
                foreach (OleObject ole in sheet.OleObjects)
                {
                    try
                    {
                        // Retrieve the binary data of the OLE object
                        byte[] oleData = ole.FullObjectBin;

                        if (oleData != null && oleData.Length > 0)
                        {
                            string outputPath = $"OleObject_{oleCounter}_Data.bin";
                            File.WriteAllBytes(outputPath, oleData);
                            Console.WriteLine($"Successfully extracted OLE object #{oleCounter} to '{outputPath}'.");
                        }
                        else
                        {
                            Console.WriteLine($"OLE object #{oleCounter} contains no data.");
                        }
                    }
                    catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted)
                    {
                        // Handle corrupted OLE object data
                        Console.WriteLine($"Error: OLE object #{oleCounter} is corrupted. Details: {ex.Message}");
                    }
                    catch (CellsException ex) when (ex.Code == ExceptionType.UnsupportedFeature)
                    {
                        // Handle unsupported OLE object formats
                        Console.WriteLine($"Error: OLE object #{oleCounter} format is unsupported. Details: {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        // General fallback for any other unexpected errors
                        Console.WriteLine($"Unexpected error while processing OLE object #{oleCounter}: {ex.Message}");
                    }

                    oleCounter++;
                }
            }

            // Save the workbook (if any modifications were made)
            try
            {
                workbook.Save(outputWorkbookPath);
                Console.WriteLine($"Workbook saved as '{outputWorkbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }

            Console.WriteLine("Processing completed.");
        }
    }
}
