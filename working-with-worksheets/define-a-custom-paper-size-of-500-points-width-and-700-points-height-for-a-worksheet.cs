// Title: Define a 500 pt × 700 pt custom paper size for a worksheet using Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, converts 500 pt width and 700 pt height to inches (1 pt = 1/72 in), applies the dimensions with Worksheet.PageSetup.CustomPaperSize, and saves the workbook as CustomPaperSize.xlsx.
// Keywords: Aspose.Cells custom paper size | Worksheet PageSetup.CustomPaperSize | C# set page size points | 500 pt by 700 pt Excel | convert points to inches Aspose
// Common Searches: Aspose.Cells set custom paper size 500 pt | C# convert points to inches for Excel page setup | PageSetup.CustomPaperSize example | how to define non‑standard page size in Aspose.Cells
// Developer Intent: Apply a 500 pt × 700 pt custom paper size to a worksheet and save the workbook.
// Use Cases: Printing labels or tickets that require exact dimensions | Generating PDFs with a non‑standard page size for marketing materials | Creating Excel templates that match a predefined physical form
// AI Prompts: Generate code to set a custom paper size using points without manual conversion. | Show how to read the current custom paper size from a worksheet. | Explain how to apply the same custom paper size to every worksheet in a workbook.

using System;
using Aspose.Cells;

// Creates a new Workbook, converts 500 pt width and 700 pt height to inches (1 pt = 1/72 in), applies the dimensions with Worksheet.PageSetup.CustomPaperSize, and saves the workbook as CustomPaperSize.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Convert points to inches (1 point = 1/72 inch)
        double widthInInches = 500.0 / 72.0;
        double heightInInches = 700.0 / 72.0;

        // Set custom paper size using the PageSetup.CustomPaperSize method (feature rule)
        worksheet.PageSetup.CustomPaperSize(widthInInches, heightInInches);

        // Save the workbook (lifecycle save rule)
        workbook.Save("CustomPaperSize.xlsx");
    }
}
