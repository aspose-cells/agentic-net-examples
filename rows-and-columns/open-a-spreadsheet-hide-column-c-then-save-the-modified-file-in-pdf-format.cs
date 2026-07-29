// Title: Hide Column C in Excel and Export to PDF with Aspose.Cells for .NET
// Description: Load an Excel workbook with Aspose.Cells, hide column C on the first worksheet, and save the result directly as a PDF using C#.
// Keywords: Aspose.Cells hide column | C# hide Excel column | Excel to PDF Aspose.Cells | hide column before PDF conversion | Aspose.Cells column visibility | export hidden column Excel to PDF
// Common Searches: how to hide a column in Excel using Aspose.Cells C# | Aspose.Cells hide column C and save as PDF | hide Excel column programmatically and export PDF | C# hide column then convert workbook to PDF | Aspose.Cells column visibility PDF export
// Developer Intent: Programmatically hide column C in an Excel sheet and generate a PDF file.
// Use Cases: Create a printable PDF report that excludes confidential data in column C. | Prepare a clean client‑facing PDF version of a spreadsheet by removing an unwanted column. | Automate batch processing to hide specific columns across multiple workbooks before PDF conversion.
// AI Prompts: Generate a C# method that receives an Excel file path, a column index to hide, and an output PDF path, using Aspose.Cells to perform the hide and conversion. | Provide sample code to hide column D in a worksheet and export the workbook to PDF with Aspose.Cells for .NET. | Write a script that iterates through a folder of .xlsx files, hides column C in each workbook, and saves each as a PDF.

using System;
using Aspose.Cells;

namespace HideColumnAndConvertToPdf
{
    // Load an Excel workbook with Aspose.Cells, hide column C on the first worksheet, and save the result directly as a PDF using C#.
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel file
            string inputFile = "input.xlsx";

            // Path for the resulting PDF file
            string outputFile = "output.pdf";

            // Load the workbook from the input file (lifecycle: load)
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide column C (zero‑based index 2)
            worksheet.Cells.HideColumn(2);

            // Save the modified workbook as PDF (lifecycle: save)
            workbook.Save(outputFile, SaveFormat.Pdf);

            Console.WriteLine("Column C hidden and workbook saved as PDF.");
        }
    }
}
