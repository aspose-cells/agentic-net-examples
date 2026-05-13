using System;
using Aspose.Cells;

namespace AsposeCellsMhtExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (can be any supported Excel format)
            string sourcePath = "input.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options specifying MHTML format
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);

            // Export only the active worksheet (single sheet)
            saveOptions.ExportActiveWorksheetOnly = true;

            // Save the HTML/MHTML as a single file
            saveOptions.SaveAsSingleFile = true;

            // Optional: improve visual presentation
            // saveOptions.PresentationPreference = true;

            // Path for the resulting MHTML file
            string outputPath = "single_sheet.mht";

            // Save the workbook as a single‑worksheet MHTML file
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved as single‑sheet MHTML to: {outputPath}");
        }
    }
}