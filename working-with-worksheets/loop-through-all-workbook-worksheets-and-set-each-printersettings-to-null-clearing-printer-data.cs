// Title: C# – Loop through all worksheets and clear PageSetup.PrinterSettings with Aspose.Cells
// Description: Load an Excel file using Aspose.Cells for .NET, iterate over each Worksheet, assign null to the PageSetup.PrinterSettings property to erase printer‑specific metadata, and save the modified workbook. This removes printer configuration from every sheet, making the file safe for distribution, archiving, or templating.
// Keywords: Aspose.Cells | C# | .NET | clear printer settings | remove printer metadata | PageSetup.PrinterSettings | Excel workbook | loop through worksheets | printer configuration | save workbook
// Common Searches: Aspose.Cells clear printer settings C# | remove printer data from Excel workbook using Aspose.Cells | set PageSetup.PrinterSettings to null for all sheets | how to erase printer metadata in .NET Excel file | loop through worksheets and reset printer configuration Aspose
// Developer Intent: Strip printer configuration from every worksheet in an Excel workbook via Aspose.Cells.
// Use Cases: Prepare a workbook for public sharing without exposing printer‑specific information. | Archive Excel files while ensuring no hidden printer references remain. | Create a clean template where all sheets start with default printing settings.
// AI Prompts: Generate C# code that uses Aspose.Cells to clear printer settings on all worksheets and save the file. | Explain why resetting PageSetup.PrinterSettings to null is important before distributing an Excel workbook. | Suggest an alternative method to remove printer metadata from a workbook without manually iterating each worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsPrinterSettingsClear
{
    // Load an Excel file using Aspose.Cells for .NET, iterate over each Worksheet, assign null to the PageSetup.PrinterSettings property to erase printer‑specific metadata, and save the modified workbook. This removes printer configuration from every sheet, making the file safe for distribution, archiving, or templating.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (replace with actual path if needed)
            string inputPath = "input.xlsx";
            // Path for the resulting workbook
            string outputPath = "output.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear printer settings by assigning null to the byte[] property
                sheet.PageSetup.PrinterSettings = null;
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
