// Title: How to save a modified Excel workbook while preserving original formatting and textures using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells save workbook after editing a cell without losing cell styles | preserve original Excel formatting when saving with Aspose.Cells .NET | keep textures and embedded objects after modifying a worksheet using Aspose.Cells | save workbook to a new file while retaining all formatting with Aspose.Cells | load and save Excel file with Aspose.Cells without affecting existing styles
// Tags: Aspose.Cells save workbook preserving formatting | C# modify cell value without altering styles Aspose.Cells | Excel Xlsx save options retain textures Aspose.Cells | load workbook default options keep original content Aspose.Cells | output new workbook path Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSavePreserveFormatting
{
    // The example loads an existing Excel file with Aspose.Cells, updates the value of cell A1 on the first worksheet, and then saves the workbook to a new file using SaveFormat.Xlsx. The default loading and saving behavior automatically retains all original formatting, styles, textures, and embedded objects, ensuring the output file looks identical to the source except for the changed cell value.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths.
            // If arguments are supplied, use them; otherwise fall back to default names in the executable directory.
            string inputPath;
            string outputPath;

            if (args.Length >= 2)
            {
                inputPath = args[0];
                outputPath = args[1];
            }
            else
            {
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                inputPath = Path.Combine(exeDir, "input.xlsx");
                outputPath = Path.Combine(exeDir, "output_preserved.xlsx");
            }

            // Validate that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Error: Input file not found at \"{inputPath}\".");
                return;
            }

            try
            {
                // Load the workbook. Using the default LoadOptions preserves all original content.
                Workbook workbook = new Workbook(inputPath);

                // Example modification: change the value of cell A1 on the first worksheet.
                // This does NOT affect any existing formatting, styles, or textures applied to the cell.
                Worksheet firstSheet = workbook.Worksheets[0];
                firstSheet.Cells["A1"].PutValue("Modified by Aspose.Cells");

                // Save the workbook to a new file while preserving all original formatting, styles, and embedded objects.
                // The Save method with a file path automatically retains formatting.
                // If you need explicit control, you can use OoxmlSaveOptions, but the default behavior is sufficient.
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook successfully saved to \"{outputPath}\" with original formatting preserved.");
            }
            catch (Exception ex)
            {
                // Report any unexpected errors without swallowing them.
                Console.Error.WriteLine("An error occurred while processing the workbook:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
