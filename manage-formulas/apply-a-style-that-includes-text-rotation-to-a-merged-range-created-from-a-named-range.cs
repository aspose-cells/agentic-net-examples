using System;
using System.IO;
using Aspose.Cells;
using ARange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

namespace AsposeCellsExamples
{
    public class ApplyRotationToMergedNamedRange
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Define the address of the range to be named and merged
                const string rangeAddress = "B2:D4";

                // Create the range
                ARange range = worksheet.Cells.CreateRange(rangeAddress);

                // Assign a name to the range (named range)
                range.Name = "MyMergedRange";

                // Merge the cells in the range
                range.Merge();

                // Put a sample value into the merged cell (top‑left cell of the range)
                worksheet.Cells["B2"].PutValue("Rotated Text");

                // Create a style with a rotation angle
                Style rotationStyle = workbook.CreateStyle();
                rotationStyle.RotationAngle = 45; // Rotate text 45 degrees

                // Enable the rotation flag so the rotation is applied
                var flag = new StyleFlag();
                flag.Rotation = true;

                // Apply the style with the rotation flag to the merged range
                range.ApplyStyle(rotationStyle, flag);

                // Define output file path
                const string outputPath = "MergedNamedRangeWithRotation.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (overwrite if it already exists)
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Log any runtime errors
                Console.WriteLine($"Runtime error: {ex.Message}");
                throw; // Re‑throw to be caught by Main if needed
            }
        }
    }
}