using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ExcelSmartArtProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output directories
            string inputDirectory = @"C:\InputExcelFiles";
            string outputDirectory = @"C:\OutputExcelFiles";

            try
            {
                // Verify input directory exists
                if (!Directory.Exists(inputDirectory))
                {
                    Console.WriteLine($"Input directory not found: {inputDirectory}");
                    return;
                }

                // Ensure output directory exists
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // Process each Excel file in the input directory
                foreach (string filePath in Directory.GetFiles(inputDirectory, "*.xlsx"))
                {
                    // Verify the file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(filePath);

                        // Iterate through all worksheets
                        foreach (Worksheet worksheet in workbook.Worksheets)
                        {
                            // Iterate through all shapes in the worksheet
                            foreach (Shape shape in worksheet.Shapes)
                            {
                                // Check if the shape is a SmartArt shape
                                if (shape.IsSmartArt)
                                {
                                    // Convert SmartArt to a GroupShape
                                    GroupShape groupShape = shape.GetResultOfSmartArt();

                                    // Optional: manipulate the resulting groupShape here
                                    // e.g., groupShape.Left += 10;
                                }
                            }
                        }

                        // Prepare save options to ensure SmartArt conversion is persisted
                        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                        {
                            UpdateSmartArt = true
                        };

                        // Save the modified workbook to the output directory
                        string outputPath = Path.Combine(outputDirectory, Path.GetFileName(filePath));
                        workbook.Save(outputPath, saveOptions);
                        Console.WriteLine($"Processed and saved: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
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