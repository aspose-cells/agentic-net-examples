// Title: C# – Load FODS, Change Default Font, and Save as ODS with Aspose.Cells
// Description: Shows how to open a FODS workbook using Aspose.Cells for .NET, set the workbook's default font (e.g., Calibri 11), apply OdsSaveOptions (LibreOffice generator), and export the file as an ODS spreadsheet.
// Keywords: Aspose.Cells | C# | FODS to ODS conversion | default font | Workbook.DefaultStyle | OdsLoadOptions | OdsSaveOptions | LibreOffice generator | OpenDocument Spreadsheet | programmatic style change
// Common Searches: Aspose.Cells change default font before saving | Convert FODS to ODS C# Aspose | Set workbook default style Aspose.Cells .NET | OdsLoadOptions example code | How to use OdsSaveOptions LibreOffice generator | C# code to modify default font in FODS
// Developer Intent: Load a .fods file, modify its default font, and save the updated workbook as a .ods document using Aspose.Cells for .NET.
// Use Cases: Standardize corporate fonts across all worksheets before exporting a FODS template to ODS for consistent appearance in LibreOffice. | Batch‑process a folder of FODS files, applying a company‑wide default font and generating ODS files with the LibreOffice generator. | Create ODS reports from FODS sources while ensuring the default style matches branding guidelines without manually editing each file.
// AI Prompts: Generate C# code that loads a FODS workbook with Aspose.Cells, sets the default font to Arial 10, and saves it as ODS using the default generator. | Explain step‑by‑step how to apply a custom Workbook.DefaultStyle to an entire workbook before exporting to ODS, including how to select a specific OdsGeneratorType. | Write a C# script that iterates through a directory of .fods files, changes each workbook's default font to Calibri 11, and exports them as .ods files using the LibreOffice generator.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsFodsToOds
{
    // Shows how to open a FODS workbook using Aspose.Cells for .NET, set the workbook's default font (e.g., Calibri 11), apply OdsSaveOptions (LibreOffice generator), and export the file as an ODS spreadsheet.
    class Program
    {
        static void Main()
        {
            // Path to the source FODS file
            string sourcePath = "input.fods";

            // Load the FODS workbook using OdsLoadOptions
            OdsLoadOptions loadOptions = new OdsLoadOptions();
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Change the default font for the entire workbook
            workbook.DefaultStyle.Font.Name = "Calibri";
            workbook.DefaultStyle.Font.Size = 11;

            // Prepare ODS save options (optional: set generator type)
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

            // Save the workbook as an ODS file
            string outputPath = "output.ods";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook converted from FODS to ODS and saved to '{outputPath}'.");
        }
    }
}
