// Title: Convert an Aspose.Cells workbook to PDF with left‑to‑right worksheet orientation in C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, disables right‑to‑left display, and saves it as a PDF. | Show how to set DisplayRightToLeft to false on a worksheet before exporting to PDF using Aspose.Cells. | Provide a minimal Aspose.Cells example that adds data, configures left‑to‑right layout, and generates a PDF file.
// Common Searches: Aspose.Cells C# export workbook to PDF with left to right layout | disable right to left view in Aspose.Cells before PDF conversion | how to set worksheet direction for western scripts in Aspose.Cells PDF output | C# save Excel file as PDF using Aspose.Cells with left‑to‑right orientation
// Tags: Aspose.Cells set DisplayRightToLeft false | Aspose.Cells PDF export left-to-right orientation | C# workbook to PDF conversion Aspose.Cells | worksheet direction western script Aspose.Cells | SaveFormat.Pdf usage Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsPdfConversion
{
    // The program creates a new Aspose.Cells workbook, ensures the first worksheet displays left‑to‑right by setting DisplayRightToLeft to false, adds a sample cell value, and saves the workbook as a PDF file named output.pdf using SaveFormat.Pdf.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the worksheet to display left‑to‑right (default for Western scripts)
            worksheet.DisplayRightToLeft = false;

            // Add sample data (optional)
            worksheet.Cells["A1"].PutValue("Hello, world!");

            // Save the workbook as PDF using the Save method with SaveFormat.Pdf
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
