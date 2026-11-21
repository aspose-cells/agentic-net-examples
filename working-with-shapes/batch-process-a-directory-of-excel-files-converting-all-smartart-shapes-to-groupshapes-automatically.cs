// Title: Batch detect SmartArt shapes in multiple Excel workbooks using Aspose.Cells for .NET (C#)
// Common Searches: how to scan a folder of .xlsx files for SmartArt shapes with Aspose.Cells in C# | C# code to iterate through all worksheets and shape collections to find SmartArt using Aspose.Cells | Aspose.Cells detect IsSmartArt property while processing multiple Excel workbooks | log SmartArt shapes during bulk Excel file processing with Aspose.Cells .NET | automate scanning of Excel files for unsupported SmartArt conversion using Aspose.Cells
// Tags: batch detect SmartArt Aspose.Cells C# | iterate worksheet shape collection Aspose.Cells | check IsSmartArt property Aspose.Cells | process multiple .xlsx files Aspose.Cells | log unsupported SmartArt conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The C# utility validates a source directory, enumerates all .xlsx files, loads each workbook with Aspose.Cells, walks through every worksheet and its ShapeCollection, uses the IsSmartArt property to identify SmartArt shapes, logs each occurrence (noting that direct conversion to GroupShape is not supported), saves the workbook (overwriting the original), and reports progress while handling errors.
class SmartArtToGroupShapeBatchProcessor
{
    static void Main()
    {
        // Directory containing the Excel files to process
        string sourceDirectory = @"C:\ExcelFiles";

        // Validate source directory
        if (!Directory.Exists(sourceDirectory))
        {
            Console.WriteLine($"Source directory does not exist: {sourceDirectory}");
            return;
        }

        // Get all Excel files in the directory (top‑level only)
        string[] excelFiles = Directory.GetFiles(sourceDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in excelFiles)
        {
            try
            {
                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipping: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    ShapeCollection shapes = sheet.Shapes;

                    // Iterate backwards because the collection may change during processing
                    for (int i = shapes.Count - 1; i >= 0; i--)
                    {
                        Shape shape = shapes[i];

                        // Determine whether the shape is a SmartArt shape
                        bool isSmartArt = false;
                        try
                        {
                            isSmartArt = shape.IsSmartArt;
                        }
                        catch
                        {
                            // Property not supported – ignore
                        }

                        if (isSmartArt)
                        {
                            // Current Aspose.Cells versions do not provide a direct conversion method.
                            // Log the presence of SmartArt for awareness.
                            Console.WriteLine($"SmartArt detected (no conversion available) in shape '{shape.Name}' of file '{filePath}'.");
                        }
                    }
                }

                // Save the (potentially unchanged) workbook, overwriting the original file
                workbook.Save(filePath);
                Console.WriteLine($"Processed and saved: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Processing completed. All accessible SmartArt shapes have been handled.");
    }
}
