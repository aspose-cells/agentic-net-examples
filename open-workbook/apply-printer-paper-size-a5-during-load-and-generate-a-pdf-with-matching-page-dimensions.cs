// Title: Set A5 Paper Size on Load and Export Excel to PDF with Aspose.Cells for .NET
// Description: Shows how to configure Aspose.Cells LoadOptions to apply the A5 printer paper size when loading an Excel workbook, optionally enforce the size on a worksheet, and save the workbook as a PDF whose pages exactly match A5 dimensions.
// Keywords: Aspose.Cells | C# | LoadOptions | SetPaperSize | A5 paper size | Excel to PDF | PageSetup | printer settings | document conversion | PDF page dimensions
// Common Searches: Aspose.Cells set A5 paper size on load | Export Excel to A5 PDF using C# | LoadOptions SetPaperSize example | How to change worksheet page size before PDF export | C# code for A5 PDF output with Aspose.Cells
// Developer Intent: Apply the A5 printer paper size during workbook loading and generate a PDF that uses the same page dimensions.
// Use Cases: Load an existing spreadsheet and ensure all printed pages use A5 without manually adjusting each sheet. | Override a specific worksheet’s page setup to A5 when the default size differs. | Create A5‑sized PDF reports from Excel files for consistent printing or distribution.
// AI Prompts: Provide C# code that loads an Excel file with A5 paper size using Aspose.Cells and saves it as a PDF. | Explain the effect of LoadOptions.SetPaperSize on workbook printing and PDF export. | Show how to set PageSetup.PaperSize to A5 for a single worksheet before converting to PDF.

using System;
using Aspose.Cells;

namespace AsposeCellsA5PdfDemo
{
    // Shows how to configure Aspose.Cells LoadOptions to apply the A5 printer paper size when loading an Excel workbook, optionally enforce the size on a worksheet, and save the workbook as a PDF whose pages exactly match A5 dimensions.
    class Program
    {
        static void Main()
        {
            // Prepare load options and set the default printer paper size to A5
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.SetPaperSize(PaperSizeType.PaperA5);

            // Load an existing workbook (replace with your actual file path)
            // The paper size defined in loadOptions will be applied to the workbook settings
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Ensure the first worksheet also uses A5 (optional, usually inherited from settings)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;

            // Save the workbook as PDF; the page dimensions will match A5 size
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
