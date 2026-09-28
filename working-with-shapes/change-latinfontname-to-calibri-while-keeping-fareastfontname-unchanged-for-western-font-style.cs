// Title: How to change the default Latin (Western) font to Calibri while preserving the FarEast font in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets the workbook's default Latin font to Calibri, leaves the FarEast font unchanged, and saves the updated file. | Write a C# method that updates only the Font.Name property of the default style's Latin font to "Calibri" without modifying the FarEastFontName, using Aspose.Cells. | Create a script that checks for the input workbook, changes the default Western font to Calibri via Aspose.Cells, and ensures the FarEast font settings remain intact before saving.
// Common Searches: Aspose.Cells set default Latin font to Calibri C# | preserve FarEastFontName when changing workbook default font Aspose.Cells | change only western font in Excel file using Aspose.Cells .NET | C# Aspose.Cells default style font name Calibri without affecting Asian fonts | how to modify default style Latin font in Aspose.Cells workbook
// Tags: Aspose.Cells default style font change | set Latin font name Calibri Aspose.Cells | preserve FarEastFontName Aspose.Cells | C# modify workbook default font | Excel workbook western font Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook with Aspose.Cells, retrieves the default style, changes the Latin (Western) font name to "Calibri" while leaving the FarEast font unchanged, reassigns the modified style, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the default style
                Style defaultStyle = workbook.DefaultStyle;

                // Change the Latin (Western) font to Calibri
                defaultStyle.Font.Name = "Calibri";

                // Apply the modified style back to the workbook
                workbook.DefaultStyle = defaultStyle;

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
