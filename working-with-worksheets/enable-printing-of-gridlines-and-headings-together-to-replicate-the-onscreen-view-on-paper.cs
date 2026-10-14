// Title: Print Excel gridlines and row/column headings together with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that configures Aspose.Cells PageSetup to print both gridlines and row/column headings, then saves the workbook as XLSX. | Show how to enable PrintGridlines and PrintHeadings properties in Aspose.Cells before exporting the worksheet to PDF. | Adapt an existing Aspose.Cells workbook to include on‑paper gridlines and headings, demonstrating the required PageSetup settings.
// Common Searches: Aspose.Cells C# enable printing of gridlines and headings in the same printout | How to print row and column headings with gridlines using Aspose.Cells .NET | C# Aspose.Cells PageSetup PrintGridlines true PrintHeadings true example | Export Excel worksheet with visible gridlines and headings using Aspose.Cells | Print layout settings for gridlines and headings in Aspose.Cells workbook
// Tags: Aspose.Cells PageSetup PrintGridlines | Aspose.Cells PageSetup PrintHeadings | C# export Excel with gridlines and headings | Aspose.Cells print layout configuration | save workbook with gridlines Aspose.Cells

using Aspose.Cells;
using System;

// // Creates a workbook, adds sample data, enables PageSetup.PrintGridlines and PrintHeadings, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data to illustrate gridlines and headings
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Enable printing of gridlines
        sheet.PageSetup.PrintGridlines = true;

        // Enable printing of row and column headings (A, B, 1, 2, …)
        sheet.PageSetup.PrintHeadings = true;

        // Save the workbook (you can also save as PDF to see the print layout)
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}
