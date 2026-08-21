// Title: C# – Apply a Custom Theme and Export the First Worksheet as PDF with Aspose.Cells
// Description: Load an XLSX workbook, define a 12‑color custom theme, apply it, hide every sheet except the first, and save that sheet as a PDF preview using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | custom theme | Excel to PDF | hide worksheets | C# PDF preview | Workbook.CustomTheme | first sheet export | SaveFormat.Pdf
// Common Searches: Aspose.Cells apply custom theme C# | Export first worksheet to PDF Aspose.Cells | Hide worksheets before PDF conversion Aspose.Cells | Create PDF preview of Excel sheet .NET | Set custom theme colors programmatically Aspose.Cells
// Developer Intent: Apply a corporate color theme to an Excel workbook and generate a PDF preview of only the first worksheet.
// Use Cases: Produce branded PDF reports by applying a company‑wide color palette and exporting the summary sheet. | Create lightweight preview PDFs for large workbooks by hiding non‑essential sheets before conversion. | Automate batch processing to enforce a standard theme across files and deliver first‑page PDFs for document management systems.
// AI Prompts: Show how to keep the original workbook unchanged while generating a themed PDF preview. | Provide code that writes the PDF preview to a MemoryStream instead of a file. | Explain how to read back the custom theme colors from a workbook for validation.

using System.Drawing;
using Aspose.Cells;

// Load an XLSX workbook, define a 12‑color custom theme, apply it, hide every sheet except the first, and save that sheet as a PDF preview using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Define 12 custom theme colors (Background1, Text1, Background2, Text2, Accent1‑Accent6, Hyperlink, Followed Hyperlink)
        Color[] customColors = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1
            Color.FromArgb(0, 0, 0),       // Text1
            Color.FromArgb(240, 240, 240), // Background2
            Color.FromArgb(50, 50, 50),    // Text2
            Color.FromArgb(0, 112, 192),   // Accent1
            Color.FromArgb(255, 192, 0),   // Accent2
            Color.FromArgb(112, 48, 160),  // Accent3
            Color.FromArgb(0, 176, 80),    // Accent4
            Color.FromArgb(255, 0, 0),     // Accent5
            Color.FromArgb(0, 176, 240),   // Accent6
            Color.FromArgb(0, 0, 255),     // Hyperlink
            Color.FromArgb(128, 0, 128)    // Followed Hyperlink
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("MyCustomTheme", customColors);

        // Hide all worksheets except the first one to generate a preview of the first sheet only
        for (int i = 1; i < workbook.Worksheets.Count; i++)
        {
            workbook.Worksheets[i].IsVisible = false;
        }

        // Save the first sheet as a PDF file (preview)
        workbook.Save("FirstSheetPreview.pdf", SaveFormat.Pdf);
    }
}
