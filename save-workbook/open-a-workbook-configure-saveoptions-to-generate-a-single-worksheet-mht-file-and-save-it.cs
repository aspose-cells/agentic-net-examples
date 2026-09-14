// Title: Save only the active worksheet as a single‑page MHT file with Aspose.Cells in C#
// AI Prompts: Write C# code that loads an Excel workbook using Aspose.Cells and exports the current sheet to a one‑page MHTML document. | Show how to configure HtmlSaveOptions so that Aspose.Cells creates a single MHT file containing just the active worksheet.
// Common Searches: Aspose.Cells C# export active sheet to MHTML as one file | How to generate a single‑page MHT from Excel using Aspose.Cells .NET | C# save only selected worksheet as MHT with Aspose.Cells HtmlSaveOptions
// Tags: Aspose.Cells HtmlSaveOptions SaveAsSingleFile | C# export active worksheet to MHT | Aspose.Cells single worksheet MHTML conversion | MHT single file generation with Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsMhtExample
{
    // The example opens 'input.xlsx' with Aspose.Cells, sets HtmlSaveOptions to MHtml, enables SaveAsSingleFile and ExportActiveWorksheetOnly (with optional PresentationPreference), and saves the active sheet as a single‑page MHT file named 'output.mht'.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (can be any supported Excel format)
            string sourcePath = "input.xlsx";

            // Open the workbook using the constructor that accepts a file name
            Workbook workbook = new Workbook(sourcePath);

            // Create HtmlSaveOptions for MHTML format
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);

            // Configure options to generate a single‑worksheet MHT file
            saveOptions.SaveAsSingleFile = true;               // Save as a single file
            saveOptions.ExportActiveWorksheetOnly = true;      // Export only the active sheet
            // Optional: improve visual presentation
            saveOptions.PresentationPreference = true;

            // Save the workbook as MHTML
            string outputPath = "output.mht";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved as single‑worksheet MHT to: {outputPath}");
        }
    }
}
