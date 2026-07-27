using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorkbookLoadSaveWithUsingDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load an existing workbook inside a using block.
                // The using statement ensures Dispose() is called automatically.
                using (Workbook workbook = new Workbook(inputPath))
                {
                    // Example modification: write a value to cell A1 of the first worksheet.
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Processed");

                    // Save the workbook to a new file.
                    // The Save(string, SaveFormat) overload is used as per the provided rule.
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                } // workbook.Dispose() is invoked here.

                Console.WriteLine($"Workbook processed and saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}