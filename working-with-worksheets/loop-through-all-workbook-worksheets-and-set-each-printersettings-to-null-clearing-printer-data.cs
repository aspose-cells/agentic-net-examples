// Title: C# – Clear Printer Settings for All Worksheets in an Aspose.Cells Workbook
// Description: Loads an Excel file with Aspose.Cells, iterates through every worksheet, sets each sheet's PageSetup.PrinterSettings to null to remove embedded printer data, and saves the result to a new file.
// Keywords: Aspose.Cells | C# | clear printer settings | PageSetup.PrinterSettings | reset printer configuration | remove printer data | Excel workbook cleanup
// Common Searches: Aspose.Cells clear printer settings C# | remove printer configuration from all worksheets | set PageSetup.PrinterSettings to null | reset printer data before saving Excel file | how to delete printer settings in Aspose.Cells
// Developer Intent: Remove printer‑specific configuration from every worksheet in a workbook.
// Use Cases: Distribute a workbook without exposing local printer setups. | Create a clean template that users can configure on their own machines. | Prevent printing inconsistencies when the file is opened on different devices.
// AI Prompts: Write C# code using Aspose.Cells that loops through all worksheets, clears the printer settings, and saves the workbook. | Show an example that logs each worksheet name while clearing its printer settings and handles possible null references. | Provide a snippet that clears printer settings and also resets page margins for every sheet in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsPrinterSettingsClear
{
    // Loads an Excel file with Aspose.Cells, iterates through every worksheet, sets each sheet's PageSetup.PrinterSettings to null to remove embedded printer data, and saves the result to a new file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Loop through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear the printer settings by setting the byte array to null
                sheet.PageSetup.PrinterSettings = null;
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
