using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtValidation
{
    public class SmartArtConversionValidator
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Process each worksheet and its shapes
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                foreach (Shape shape in worksheet.Shapes)
                {
                    if (shape.IsSmartArt)
                    {
                        // Convert SmartArt to GroupShape
                        GroupShape groupShape = shape.GetResultOfSmartArt();

                        if (groupShape != null)
                        {
                            // Example: reposition the group shape
                            groupShape.Left = 300;
                            groupShape.Top = 100;
                        }
                        else
                        {
                            Console.WriteLine($"SmartArt shape at Z-order {shape.ZOrderPosition} could not be converted.");
                        }
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}