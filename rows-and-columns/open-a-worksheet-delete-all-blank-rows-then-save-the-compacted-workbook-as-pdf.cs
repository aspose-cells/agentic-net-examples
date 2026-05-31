using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class DeleteBlankRowsAndSavePdf
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or any specific worksheet you need)
                Worksheet worksheet = workbook.Worksheets[0];

                // Delete all blank rows in the worksheet
                worksheet.Cells.DeleteBlankRows();

                // Save the compacted workbook as a PDF file
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"Successfully saved PDF to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}