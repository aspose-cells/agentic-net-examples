using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    class SaveWorkbookWithSlicers
    {
        static void Main()
        {
            try
            {
                string inputPath = "input_with_slicers.xlsx";
                string outputPath = "output_excel2016.xlsx";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook that contains slicers
                Workbook workbook = new Workbook(inputPath);

                // Set compliance to preserve slicers (default is sufficient)
                workbook.Settings.Compliance = OoxmlCompliance.Ecma376_2006;

                // Save the workbook preserving all advanced features
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}