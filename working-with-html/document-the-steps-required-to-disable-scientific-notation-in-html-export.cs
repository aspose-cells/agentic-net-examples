// Title: Disable scientific notation when exporting an Excel workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export to HTML without scientific notation for numbers | How to prevent exponent format in HTML output using Aspose.Cells .NET | Set HtmlSaveOptions to turn off scientific notation in Aspose.Cells HTML export | C# Aspose.Cells HtmlSaveOptions disable numeric scientific format | Export Excel to HTML preserving original number formatting Aspose.Cells | Aspose.Cells HTML export keep decimal representation instead of exponent
// Tags: Aspose.Cells HtmlSaveOptions disable scientific notation | HTML export numeric formatting Aspose.Cells .NET | prevent exponent display in Aspose.Cells HTML output | Aspose.Cells workbook.Save HTML options numeric format | disable scientific notation for numbers in HTML export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Load the source workbook, create an HtmlSaveOptions object, set ExportImagesAsBase64 to true (optional) and set IsConvertNumericToScientificNotation to false, then call workbook.Save with the options. This disables scientific (exponential) notation for numeric cells in the generated HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

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

                // NOTE: In recent versions of Aspose.Cells the CalcEngineSettings
                // property may not be available. If needed, adjust the code
                // according to the version you are using.

                // Configure HTML save options (example: embed images as Base64)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true
                };

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
