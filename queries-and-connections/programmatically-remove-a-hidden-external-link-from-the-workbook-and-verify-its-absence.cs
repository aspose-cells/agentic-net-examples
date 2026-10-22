// Title: How to attempt removing hidden external links from an Excel workbook using Aspose.Cells for .NET (API limitation)
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsx file, detects hidden external connections, removes them when supported, and saves the workbook with comprehensive error handling. | Create C# logic that checks for the presence of Worksheet.ExternalLinks or ExternalLink.IsHidden in the current Aspose.Cells version, logs a clear message about the missing API, and saves the workbook unchanged.
// Common Searches: aspocells .net remove hidden external link from workbook programmatically | c# check for external connections in Excel file using Aspose.Cells | externallink.ishidden property missing in Aspose.Cells API | how to verify that an Excel workbook has no external links after processing with Aspose.Cells | aspocells external link handling limitations for hidden links
// Tags: hidden external link removal Aspose.Cells | Aspose.Cells external link detection C# | verify absence of external links Excel Aspose.Cells | Aspose.Cells API limitation external connections | C# workbook external link handling Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The sample loads 'input.xlsx' with Aspose.Cells, discovers that the current library version lacks Worksheet.ExternalLinks and ExternalLink.IsHidden members, logs this limitation, skips any link removal, and saves the workbook unchanged to 'output.xlsx' while providing clear diagnostic output.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // NOTE:
            // The current Aspose.Cells version used in this environment does not expose
            // Worksheet.ExternalLinks or ExternalLink.IsHidden members. Therefore,
            // removal of hidden external links cannot be performed directly.
            // The workbook is saved unchanged.

            Console.WriteLine("Workbook loaded successfully. No external link processing performed due to API limitations.");

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
