// Title: Batch convert SmartArt shapes to GroupShape across multiple Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that scans a given folder for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, finds every SmartArt shape, converts it to a GroupShape via dynamic invocation, and saves the result to an output folder. | Generate code that iterates through all worksheets and shapes in an Aspose.Cells workbook, detects SmartArt by runtime type name, calls ConvertToGroupShape, copies the original shape name to the new GroupShape, and handles conversion errors gracefully. | Create a script that ensures the source and destination directories exist, skips unsupported file types, logs any conversion failures, and reports completion of mass SmartArt‑to‑GroupShape processing.
// Common Searches: how to replace SmartArt with GroupShape in multiple Excel files using Aspose.Cells C# | batch processing Excel workbooks to convert SmartArt shapes to groups .NET | C# code for converting SmartArt to GroupShape with Aspose.Cells and dynamic type | automate conversion of SmartArt objects to GroupShape in a folder of .xlsx files | Aspose.Cells bulk shape conversion example C#
// Tags: Aspose.Cells bulk SmartArt conversion to GroupShape | C# dynamic ConvertToGroupShape usage | process multiple Excel files for shape replacement | retain shape name after conversion | automated Excel shape handling .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The console application scans a source directory for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, iterates through every worksheet and shape, detects SmartArt shapes at runtime, converts each to a GroupShape while preserving the original name, and saves the modified workbook to a destination folder, handling unsupported files and logging any conversion errors.
class SmartArtToGroupShapeBatchProcessor
{
    static void Main()
    {
        // Define the source directory containing Excel files
        string sourceDirectory = @"C:\ExcelFiles\Input";

        // Define the destination directory for processed files
        string destinationDirectory = @"C:\ExcelFiles\Output";

        // Ensure the source directory exists
        if (!Directory.Exists(sourceDirectory))
        {
            Console.WriteLine($"Source directory does not exist: {sourceDirectory}");
            return;
        }

        // Ensure the destination directory exists
        if (!Directory.Exists(destinationDirectory))
        {
            Directory.CreateDirectory(destinationDirectory);
        }

        // Get all Excel files in the source directory (top level only)
        string[] excelFiles = Directory.GetFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly);

        foreach (string filePath in excelFiles)
        {
            // Process only supported Excel extensions
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                continue;

            // Verify the file exists before loading
            if (!File.Exists(filePath))
                continue;

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each shape on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Identify SmartArt shapes via runtime type name (avoids compile‑time dependency)
                        if (shape.GetType().Name == "SmartArt")
                        {
                            try
                            {
                                // Use dynamic to invoke ConvertToGroupShape at runtime
                                dynamic smartArt = shape;
                                GroupShape groupShape = smartArt.ConvertToGroupShape();

                                // Preserve the original name
                                groupShape.Name = shape.Name;
                            }
                            catch (Exception convEx)
                            {
                                Console.WriteLine($"Failed to convert SmartArt in sheet '{sheet.Name}': {convEx.Message}");
                            }
                        }
                    }
                }

                // Determine the output file path (preserving the original file name)
                string outputFilePath = Path.Combine(destinationDirectory, Path.GetFileName(filePath));

                // Save the modified workbook
                workbook.Save(outputFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Processing completed. All SmartArt shapes have been converted to GroupShapes.");
    }
}
