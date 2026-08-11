// Title: Set Custom Page Margins (cm) and Freeze Top Rows with Print Titles in Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, configures page margins in centimeters using PageSetup, freezes the first three rows, marks those rows as PrintTitleRows so they repeat on every printed page, and saves the file as CustomMarginAndFreeze.xlsx.
// Keywords: Aspose.Cells C# | custom page margins centimeters | PageSetup margins | FreezePanes rows | PrintTitleRows header repeat | worksheet printable area | .NET spreadsheet margins | freeze top rows Aspose.Cells | repeat header rows printing
// Common Searches: Aspose.Cells set margins in centimeters | how to freeze rows and repeat on print in Aspose.Cells | C# PageSetup PrintTitleRows example | freeze top rows Aspose.Cells .NET | custom margin settings Aspose.Cells workbook
// Developer Intent: Apply precise margin dimensions and lock the top rows so they stay visible while scrolling and automatically repeat on each printed page.
// Use Cases: Design a printable report where column headings stay fixed and appear on every page. | Generate invoices with standardized margins and repeated header rows for consistency. | Create a corporate spreadsheet template that enforces exact margin sizes and frozen header rows for printing.
// AI Prompts: Generate C# code using Aspose.Cells to set page margins in inches, freeze the first two rows, and repeat them on each printed page. | Explain the role of PageSetup.PrintTitleRows for repeating header rows when printing a workbook with Aspose.Cells for .NET. | Provide a step‑by‑step tutorial to apply custom margins and freeze panes in an existing workbook using Aspose.Cells C#.

using System;
using Aspose.Cells;

// C# example that creates a workbook, configures page margins in centimeters using PageSetup, freezes the first three rows, marks those rows as PrintTitleRows so they repeat on every printed page, and saves the file as CustomMarginAndFreeze.xlsx.
class CustomMarginAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the PageSetup object to define custom margins (centimeters)
        PageSetup pageSetup = worksheet.PageSetup;
        pageSetup.TopMargin = 2.0;      // Top margin = 2 cm
        pageSetup.BottomMargin = 1.5;   // Bottom margin = 1.5 cm
        pageSetup.LeftMargin = 1.0;     // Left margin = 1 cm
        pageSetup.RightMargin = 1.0;    // Right margin = 1 cm
        pageSetup.HeaderMargin = 0.5;   // Header margin = 0.5 cm
        pageSetup.FooterMargin = 0.5;   // Footer margin = 0.5 cm

        // Freeze the first three rows so they stay visible and within the printable area
        // Parameters: row index, column index, number of frozen rows, number of frozen columns
        worksheet.FreezePanes(3, 0, 3, 0);

        // Ensure the frozen rows repeat on each printed page
        pageSetup.PrintTitleRows = "$1:$3";

        // Save the workbook
        workbook.Save("CustomMarginAndFreeze.xlsx");
    }
}
