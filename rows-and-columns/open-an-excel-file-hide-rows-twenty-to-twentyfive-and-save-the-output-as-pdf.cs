// Title: Hide Rows 20‑25 in Excel and Export to PDF with Aspose.Cells for .NET
// Description: Load an Excel workbook, hide rows 20‑25 using Aspose.Cells (zero‑based index), and save the modified sheet directly as a PDF file.
// Keywords: Aspose.Cells | C# | HideRows | Excel to PDF | hide rows 20-25 | Cells.HideRows | PDF export | Aspose.Cells .NET | row visibility | worksheet PDF conversion
// Common Searches: hide rows 20 to 25 Aspose.Cells | Aspose.Cells hide rows before PDF export | C# hide specific rows Excel PDF | how to hide rows in Aspose.Cells | export hidden rows Excel to PDF .NET
// Developer Intent: Hide rows 20‑25 in an Excel worksheet and generate a PDF using Aspose.Cells.
// Use Cases: Create printable reports that exclude internal calculation rows. | Generate PDF invoices while omitting draft or confidential rows. | Distribute spreadsheet data as PDF without showing hidden sections.
// AI Prompts: Provide C# code that uses Aspose.Cells to hide rows 20‑25 and save the workbook as a PDF. | Explain how Cells.HideRows works with zero‑based indexing in Aspose.Cells. | Show error‑handling patterns for missing input files when converting Excel to PDF with hidden rows.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load an Excel workbook, hide rows 20‑25 using Aspose.Cells (zero‑based index), and save the modified sheet directly as a PDF file.
    public class HideRowsAndSavePdf
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing Excel file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Hide rows 20 to 25 (1‑based). Cells.HideRows uses zero‑based index.
                worksheet.Cells.HideRows(19, 6);

                // Save the modified workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"PDF saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideRowsAndSavePdf.Run();
        }
    }
}
