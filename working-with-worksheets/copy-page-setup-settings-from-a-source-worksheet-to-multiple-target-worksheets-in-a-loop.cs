// Title: Copy PageSetup from a Source Worksheet to Multiple Worksheets with Aspose.Cells for .NET
// Description: Demonstrates how to load or create an Excel workbook, select the first worksheet as a template, and loop through the remaining sheets to copy its PageSetup settings using Aspose.Cells' CopyOptions. The modified workbook is then saved.
// Keywords: Aspose.Cells | C# | .NET | Copy PageSetup | Worksheet print settings | CopyOptions | Excel automation | loop through worksheets | duplicate page layout | transfer margins and orientation
// Common Searches: Aspose.Cells copy PageSetup to other sheets C# | How to duplicate print settings across worksheets using Aspose.Cells | Copy worksheet page layout in a loop .NET | Transfer margins and header/footer with Aspose.Cells | CopyOptions usage for PageSetup in C#
// Developer Intent: Replicate the PageSetup configuration of one worksheet across all other worksheets in the same workbook.
// Use Cases: Apply identical print orientation, margins, and scaling before exporting the workbook to PDF. | Ensure consistent header, footer, and page numbering across multiple sheets in a generated report. | Synchronize page scaling and fit‑to‑page options when programmatically adding new worksheets that must match a template.
// AI Prompts: Generate C# code with Aspose.Cells that copies the PageSetup from a template worksheet to a list of target worksheets, creating the input file if it does not exist. | Show how to copy only specific PageSetup properties such as margins and orientation using CopyOptions in Aspose.Cells. | Explain how to modify the loop to skip hidden worksheets while copying PageSetup settings with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to load or create an Excel workbook, select the first worksheet as a template, and loop through the remaining sheets to copy its PageSetup settings using Aspose.Cells' CopyOptions. The modified workbook is then saved.
    public class CopyPageSetupDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Ensure the input file exists; create a new workbook if it does not.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Add an extra worksheet to demonstrate copying.
                workbook.Worksheets.Add();
                workbook.Save(inputPath);
                Console.WriteLine($"Input file not found. Created a new workbook at '{inputPath}'.");
            }

            // Source worksheet (first sheet)
            Worksheet sourceWorksheet = workbook.Worksheets[0];

            // Default copy options
            CopyOptions copyOptions = new CopyOptions();

            // Copy page setup from source to all other worksheets
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (i == workbook.Worksheets.IndexOf(sourceWorksheet))
                    continue; // Skip the source worksheet

                Worksheet targetWorksheet = workbook.Worksheets[i];
                targetWorksheet.PageSetup.Copy(sourceWorksheet.PageSetup, copyOptions);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
