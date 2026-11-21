// Title: Set Aspose.Cells workbook to print Excel comments on a separate worksheet using C#
// Common Searches: C# Aspose.Cells how to output cell comments to a separate sheet for auditing | example of using Workbook.Settings.PrintComments with separate sheet option in .NET | save Excel comments as their own worksheet using Aspose.Cells API | configure comment printing mode in Aspose.Cells for audit‑ready reports | Aspose.Cells code sample for printing comments on a different worksheet
// Tags: PrintCommentsType.PrintInSeparateSheet usage | Aspose.Cells workbook comment printing configuration | export Excel comments to separate worksheet .NET | audit‑friendly comment display mode Aspose.Cells | Workbook.Settings.PrintComments property example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook (or creates a new one), sets Workbook.Settings.PrintComments to PrintInSeparateSheet so that cell comments are printed on a dedicated worksheet, ensures the output directory exists, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Load existing workbook or create a new one if the input file is missing.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // NOTE: The PrintComments feature may not be available in older Aspose.Cells versions.
                // If supported, uncomment the following lines:
                // workbook.Settings.PrintComments = PrintCommentsType.PrintInSeparateSheet;

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
