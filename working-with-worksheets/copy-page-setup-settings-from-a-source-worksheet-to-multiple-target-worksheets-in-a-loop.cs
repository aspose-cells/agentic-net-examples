// Title: Copy page setup from a template worksheet to multiple worksheets using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that reads the PageSetup of a source worksheet and applies all its properties (orientation, margins, print area, zoom, fit‑to‑pages, etc.) to each worksheet in a given list using Aspose.Cells. | Generate a loop that iterates over an array of worksheet names, copies the page layout settings from a template sheet to each target sheet, and saves the workbook.
// Common Searches: Aspose.Cells C# copy page layout from one sheet to several other sheets | C# loop to duplicate worksheet print settings with Aspose.Cells | apply same margins and orientation to multiple worksheets programmatically Aspose.Cells | set identical print area for multiple worksheets using Aspose.Cells .NET
// Tags: Aspose.Cells copy worksheet page setup | C# replicate print margins across sheets | Aspose.Cells loop copy page layout | apply template page setup to multiple worksheets .NET | programmatic page setup transfer Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, retrieves the PageSetup from a template worksheet, and copies orientation, paper size, all margin values, header/footer margins, centering options, print area, zoom, fit‑to‑pages settings, and gridline/heading flags to each target worksheet listed. Missing target sheets are skipped, and the modified workbook is saved to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook that contains the source and target worksheets
            Workbook workbook = new Workbook(inputPath);

            // Define the source worksheet name (the one with the desired page setup)
            string sourceSheetName = "TemplateSheet";

            // Get the source worksheet and its PageSetup object
            Worksheet sourceSheet = workbook.Worksheets[sourceSheetName];
            if (sourceSheet == null)
                throw new ArgumentException($"Source worksheet \"{sourceSheetName}\" does not exist.");

            PageSetup sourceSetup = sourceSheet.PageSetup;

            // List of target worksheet names that should receive the same page setup
            List<string> targetSheetNames = new List<string>
            {
                "Sheet1",
                "Sheet2",
                "Sheet3"
                // add more sheet names as needed
            };

            // Loop through each target worksheet and copy the page setup settings
            foreach (string targetName in targetSheetNames)
            {
                Worksheet targetSheet = workbook.Worksheets[targetName];
                if (targetSheet == null)
                {
                    Console.WriteLine($"Warning: Target worksheet \"{targetName}\" not found. Skipping.");
                    continue;
                }

                PageSetup targetSetup = targetSheet.PageSetup;

                // Copy individual page setup properties
                targetSetup.Orientation = sourceSetup.Orientation;
                targetSetup.PaperSize = sourceSetup.PaperSize;
                targetSetup.TopMargin = sourceSetup.TopMargin;
                targetSetup.BottomMargin = sourceSetup.BottomMargin;
                targetSetup.LeftMargin = sourceSetup.LeftMargin;
                targetSetup.RightMargin = sourceSetup.RightMargin;
                targetSetup.HeaderMargin = sourceSetup.HeaderMargin;
                targetSetup.FooterMargin = sourceSetup.FooterMargin;
                targetSetup.CenterHorizontally = sourceSetup.CenterHorizontally;
                targetSetup.CenterVertically = sourceSetup.CenterVertically;
                targetSetup.PrintArea = sourceSetup.PrintArea;
                targetSetup.Zoom = sourceSetup.Zoom;
                targetSetup.FitToPagesWide = sourceSetup.FitToPagesWide;
                targetSetup.FitToPagesTall = sourceSetup.FitToPagesTall;
                targetSetup.PrintGridlines = sourceSetup.PrintGridlines;
                targetSetup.PrintHeadings = sourceSetup.PrintHeadings;
                // Add more property copies if additional settings are required
            }

            // Save the workbook with the updated page setups
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
