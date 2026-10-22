// Title: How to copy the first worksheet’s PageSetup to all other worksheets in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, reads the PageSetup of the first sheet, and applies it to every other sheet using a loop. | Write a method that accepts a Workbook object and transfers the source PageSetup to each target worksheet via the PageSetup.Copy function with default CopyOptions. | Create a console application that checks for an input .xlsx file, propagates its first sheet’s page layout to all remaining sheets, and saves the updated workbook.
// Common Searches: Aspose.Cells C# copy page setup from first worksheet to all worksheets in a workbook | How to duplicate Excel sheet page layout using Aspose.Cells .NET | Loop through worksheets and apply same PageSetup settings with Aspose.Cells | Copy PageSetup properties between worksheets Aspose.Cells C# example | Programmatically propagate page margins and orientation to all sheets using Aspose.Cells
// Tags: Aspose.Cells PageSetup.Copy method | C# loop replicate worksheet page layout | propagate Excel page margins with Aspose.Cells | apply first sheet page orientation to all worksheets | copy worksheet print settings programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads 'input.xlsx', retrieves the PageSetup of the first worksheet, iterates over the remaining worksheets copying those settings with PageSetup.Copy and default CopyOptions, and saves the modified workbook as 'output.xlsx' while handling missing files and exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the PageSetup of the first worksheet (source settings)
                PageSetup sourceSetup = workbook.Worksheets[0].PageSetup;

                // Copy page‑setup settings to all other worksheets
                for (int i = 1; i < workbook.Worksheets.Count; i++)
                {
                    PageSetup targetSetup = workbook.Worksheets[i].PageSetup;
                    // Use default CopyOptions to copy all properties
                    targetSetup.Copy(sourceSetup, new CopyOptions());
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
