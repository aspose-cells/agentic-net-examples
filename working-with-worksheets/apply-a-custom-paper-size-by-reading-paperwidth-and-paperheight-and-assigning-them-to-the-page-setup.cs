// Title: Set a Custom Paper Size in Aspose.Cells (.NET) Using Width and Height
// Description: Shows how to create a workbook, access its first worksheet, and assign a custom paper size by providing width and height (in inches) via PageSetup.CustomPaperSize, then save the Excel file.
// Keywords: Aspose.Cells | custom paper size | PageSetup.CustomPaperSize | .NET | C# | worksheet page setup | set paper dimensions | Excel printing | non‑standard paper size
// Common Searches: Aspose.Cells set custom paper size .NET | PageSetup.CustomPaperSize example C# | how to define worksheet page dimensions in Aspose.Cells | print Excel on custom sized paper using Aspose.Cells | custom paper size for Excel workbook Aspose
// Developer Intent: Define a worksheet’s page size with exact width and height values instead of using predefined paper sizes.
// Use Cases: Print receipts or labels that require non‑standard dimensions by configuring the worksheet before export. | Generate reports for specialized printers that demand a specific page size, ensuring accurate layout on the target media. | Create marketing flyers or custom‑shaped documents where the Excel sheet must match unique dimensions for perfect printing.
// AI Prompts: Provide a C# snippet that sets a custom paper size in Aspose.Cells using variables for width and height. | Explain how to read paper width and height from a JSON config file and apply them with PageSetup.CustomPaperSize in Aspose.Cells. | Show how to export a workbook with a custom paper size to PDF while preserving the exact dimensions.

using System;
using Aspose.Cells;

namespace CustomPaperSizeDemo
{
    // Shows how to create a workbook, access its first worksheet, and assign a custom paper size by providing width and height (in inches) via PageSetup.CustomPaperSize, then save the Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: read custom dimensions (in inches)
            double customWidth = 2.5;   // width in inches
            double customHeight = 3.5;  // height in inches

            // Set the worksheet to use a custom paper size
            sheet.PageSetup.PaperSize = PaperSizeType.Custom;
            sheet.PageSetup.CustomPaperSize(customWidth, customHeight);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("CustomPaperSizeOutput.xlsx");
        }
    }
}
