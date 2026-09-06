// Title: Set a corporate typeface for headings and body text by updating the default style font in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx file, change the workbook’s default font to a corporate typeface for headings and body, and save the changes using Aspose.Cells in C#. | Replace the theme fonts in a workbook by assigning a custom corporate font to Workbook.DefaultStyle.Font.Name and persist the file. | Programmatically apply a corporate font to all cells by updating the workbook’s default style and handling missing input files with Aspose.Cells.
// Common Searches: how to set a custom default font for an Excel workbook using Aspose.Cells C# | Aspose.Cells change workbook theme fonts to corporate typeface | C# update default style font name in existing .xlsx with Aspose.Cells | apply corporate font to headings and body text in Excel programmatically | Aspose.Cells set global font for all cells in a workbook
// Tags: Aspose.Cells modify workbook theme fonts | C# apply corporate typeface to Excel workbook | set global font for .xlsx using Aspose.Cells | programmatic font scheme change in Excel | update workbook font scheme Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads 'input.xlsx', sets Workbook.DefaultStyle.Font.Name to a corporate typeface, ensures the output directory exists, saves the modified workbook as 'output.xlsx', and logs any errors such as a missing input file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // -----------------------------------------------------------------
                // Update workbook fonts.
                // Aspose.Cells versions prior to 22.x expose Theme as a string,
                // so we modify the default style as a fallback to apply a corporate
                // typeface throughout the workbook.
                // -----------------------------------------------------------------
                workbook.DefaultStyle.Font.Name = "CorporateBodyFont";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
