using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class UnhideAndReorderWorksheet
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Input workbook path
                string inputPath = "input.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Worksheet to unhide
                string sheetNameToUnhide = "HiddenSheet";

                // Locate the worksheet by name
                Worksheet sheet = workbook.Worksheets[sheetNameToUnhide];
                if (sheet == null)
                {
                    Console.WriteLine($"Worksheet \"{sheetNameToUnhide}\" not found.");
                    return;
                }

                // Unhide the worksheet
                sheet.IsVisible = true;

                // Move worksheet to desired index (zero‑based)
                int desiredIndex = 1;
                sheet.MoveTo(desiredIndex);

                // Save the updated workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Worksheet \"{sheetNameToUnhide}\" is now visible and moved to index {desiredIndex}.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}