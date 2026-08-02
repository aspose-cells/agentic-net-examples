// Title: C# – Load Excel workbook with Aspose.Cells and log each worksheet’s Automatic Paper Size
// Description: Shows how to open an XLSX file using Aspose.Cells for .NET, iterate through all worksheets, read the PageSetup.IsAutomaticPaperSize flag, write the sheet name and value to the console, and optionally save the workbook while handling missing files and runtime errors.
// Keywords: Aspose.Cells | .NET | C# | load workbook | read IsAutomaticPaperSize | worksheet page setup | log automatic paper size | save workbook | exception handling | Excel file processing
// Common Searches: Aspose.Cells read automatic paper size | C# get worksheet IsAutomaticPaperSize property | How to check automatic paper size with Aspose.Cells | Iterate worksheets and log page setup Aspose.Cells | Save workbook after reading page setup properties
// Developer Intent: Load an existing Excel file, retrieve the IsAutomaticPaperSize flag for every worksheet, output the results, and optionally save the workbook.
// Use Cases: Audit page‑setup settings before batch printing. | Generate a console report of paper‑size configuration for compliance checks. | Identify worksheets with manual paper size to modify them programmatically. | Log workbook settings as part of an automated QA pipeline.
// AI Prompts: Provide C# code using Aspose.Cells to set IsAutomaticPaperSize = true for all worksheets in a loaded workbook. | Create a CSV report of each worksheet’s automatic paper size status with Aspose.Cells. | Show robust error handling for missing input files and permission issues when reading page‑setup properties.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to open an XLSX file using Aspose.Cells for .NET, iterate through all worksheets, read the PageSetup.IsAutomaticPaperSize flag, write the sheet name and value to the console, and optionally save the workbook while handling missing files and runtime errors.
    public class LogAutomaticPaperSize
    {
        public static void Run()
        {
            // Path to the existing Excel file
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and log the IsAutomaticPaperSize property
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    bool isAuto = sheet.PageSetup.IsAutomaticPaperSize;
                    Console.WriteLine($"Worksheet '{sheet.Name}' Automatic Paper Size: {isAuto}");
                }

                // Optionally save the workbook (demonstrates the save rule)
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during processing
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LogAutomaticPaperSize.Run();
        }
    }
}
