using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class SplitPdfPerWorksheet
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

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet and save it as a separate PDF
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];

                    // Configure PDF save options for the current sheet only
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        SheetSet = new SheetSet(new int[] { i }),
                        OnePagePerSheet = true
                    };

                    // Build output file name using the worksheet name
                    string outputFileName = $"{sheet.Name}.pdf";

                    try
                    {
                        // Save the current worksheet as PDF
                        workbook.Save(outputFileName, pdfOptions);
                        Console.WriteLine($"Saved: {outputFileName}");
                    }
                    catch (Exception saveEx)
                    {
                        Console.WriteLine($"Failed to save sheet '{sheet.Name}': {saveEx.Message}");
                    }
                }
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
            }
        }
    }
}