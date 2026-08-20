// Title: C# – Retrieve Paper Width of the First Worksheet from an XLSX Workbook with Aspose.Cells
// Description: Load an XLSX file using Aspose.Cells for .NET, access the first worksheet, read its PageSetup.PaperWidth property (in inches), and output the value to the console.
// Keywords: Aspose.Cells C# paper width | Worksheet PageSetup PaperWidth | read XLSX page setup .NET | get worksheet paper size inches | Aspose.Cells workbook page dimensions
// Common Searches: Aspose.Cells get worksheet paper width C# | how to read PageSetup PaperWidth from XLSX | C# retrieve first sheet paper size using Aspose | Aspose.Cells page setup dimensions example
// Developer Intent: Read the paper width setting of the first worksheet in an XLSX workbook.
// Use Cases: Verify that a worksheet fits standard printer paper before printing. | Calculate scaling or layout adjustments based on the sheet's paper width. | Enforce corporate printing standards by checking page‑setup dimensions.
// AI Prompts: Generate C# code with Aspose.Cells that lists the PaperWidth of every worksheet in centimeters. | Show how to modify a worksheet's PaperWidth, save the workbook, and confirm the change programmatically. | Explain how to compare the PaperWidth of multiple worksheets against a target page size and flag mismatches.

using System;
using Aspose.Cells;

namespace AsposeCellsPaperWidthDemo
{
    // Load an XLSX file using Aspose.Cells for .NET, access the first worksheet, read its PageSetup.PaperWidth property (in inches), and output the value to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing XLSX file
            string filePath = "input.xlsx";

            // Load the workbook from the file (uses the provided Workbook(string) constructor)
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet in the workbook
            Worksheet firstWorksheet = workbook.Worksheets[0];

            // Retrieve the paper width (in inches) from the worksheet's PageSetup
            double paperWidthInInches = firstWorksheet.PageSetup.PaperWidth;

            // Output the paper width
            Console.WriteLine($"Paper Width of the first worksheet: {paperWidthInInches} inches");
        }
    }
}
