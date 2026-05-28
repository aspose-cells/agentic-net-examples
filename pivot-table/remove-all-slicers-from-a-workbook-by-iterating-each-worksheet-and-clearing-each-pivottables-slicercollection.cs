using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsExamples
{
    public class RemoveAllSlicersDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook from the existing file
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet and clear slicers
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    SlicerCollection slicers = sheet.Slicers;
                    if (slicers != null && slicers.Count > 0)
                    {
                        slicers.Clear(); // Remove all slicers from this worksheet
                    }
                }

                // Save the modified workbook to a new file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}