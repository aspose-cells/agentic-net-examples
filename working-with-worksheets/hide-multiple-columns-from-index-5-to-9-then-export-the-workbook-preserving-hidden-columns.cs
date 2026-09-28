// Title: Hide columns 5‑9 in an Excel worksheet and export the workbook to PDF while preserving hidden columns with Aspose.Cells for .NET
// AI Prompts: Load an existing Excel file, conceal columns 5 through 9 on the first worksheet, and save the workbook as a PDF using Aspose.Cells for .NET so the column visibility state is retained in the PDF. | Using C#, hide a range of columns (indices 5‑9) in a worksheet, configure PdfSaveOptions, and export the workbook to PDF while keeping the columns hidden in the output.
// Common Searches: Aspose.Cells hide specific column range before converting Excel to PDF in C# | C# export Excel to PDF with hidden columns using Aspose.Cells | How to keep hidden columns invisible in PDF output with Aspose.Cells for .NET | Hide columns 5 to 9 in worksheet and save as PDF using Aspose.Cells API | PdfSaveOptions hide columns Aspose.Cells example
// Tags: column visibility Aspose.Cells C# | PDF export with column visibility control Aspose.Cells | worksheet column range conceal Aspose.Cells | PdfSaveOptions column visibility settings | Excel to PDF conversion Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads 'input.xlsx', conceals columns 5‑9 on the first worksheet, and saves the workbook as 'output.pdf' using PdfSaveOptions, which retains the hidden column state in the generated PDF.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Hide columns 5 through 9 (inclusive)
                // HideColumns(startColumn, totalColumns) hides a range of columns
                worksheet.Cells.HideColumns(5, 5); // hides columns 5,6,7,8,9

                // Prepare PDF save options; hidden columns are omitted by default.
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF
                workbook.Save(outputFile, pdfOptions);

                Console.WriteLine($"Workbook successfully saved to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
