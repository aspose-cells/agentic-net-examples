// Title: Aspose.Cells C# – Get Paper Width of the First Worksheet in an XLSX File
// Description: Loads an XLSX workbook with Aspose.Cells, accesses the first worksheet, reads its PageSetup.PaperWidth (in inches), and writes the value to the console.
// Keywords: Aspose.Cells C# PaperWidth | worksheet page setup width | retrieve paper size Aspose.Cells | XLSX print dimensions .NET | PageSetup PaperWidth property | read worksheet print settings | C# get worksheet paper width
// Common Searches: Aspose.Cells get worksheet paper width | PageSetup PaperWidth C# example | read print width of first sheet Aspose | how to obtain worksheet paper size in inches | C# Aspose.Cells retrieve page layout dimensions
// Developer Intent: Read the paper width (in inches) of the first worksheet in an XLSX workbook using Aspose.Cells for .NET.
// Use Cases: Verify that a sheet conforms to a specific paper size before printing. | Dynamically adjust layout by comparing the current width with target dimensions. | Audit workbooks to ensure they meet corporate printing standards.
// AI Prompts: Generate C# code that reads the PaperWidth of any worksheet index and returns the value in centimeters. | Create a snippet that checks PaperWidth and sets a custom page size when the width exceeds 8.5 inches. | Write a reusable method to fetch the paper width for a given worksheet and handle missing PageSetup settings.

using System;
using Aspose.Cells;

namespace AsposeCellsPaperWidthDemo
{
    // Loads an XLSX workbook with Aspose.Cells, accesses the first worksheet, reads its PageSetup.PaperWidth (in inches), and writes the value to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing XLSX file
            string filePath = "input.xlsx";

            // Load the workbook using the provided constructor (lifecycle rule)
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (index 0)
            Worksheet firstSheet = workbook.Worksheets[0];

            // Retrieve the paper width (in inches) from the worksheet's PageSetup
            double paperWidthInInches = firstSheet.PageSetup.PaperWidth;

            // Output the result
            Console.WriteLine($"Paper Width of the first worksheet: {paperWidthInInches} inches");
        }
    }
}
