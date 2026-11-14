// Title: How to unhide row 12, set its height to 20 points, and export an Excel workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# unhide a specific row and set row height before saving as PDF | C# code to make row 12 visible and change its height to 20 points with Aspose.Cells | Export Excel to PDF with modified row visibility using Aspose.Cells .NET | Set row height and unhide rows in Aspose.Cells then convert workbook to PDF | How to adjust row properties in Aspose.Cells and generate a PDF file
// Tags: unhide row Aspose.Cells C# | set row height Aspose.Cells | Aspose.Cells export workbook to PDF | modify row visibility before PDF conversion | adjust row properties Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file, unhides row 12, sets its height to 20 points, and saves the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Row 12 corresponds to zero‑based index 11
            int rowIndex = 11;

            // Retrieve the row via Cells.Rows collection
            Row row = sheet.Cells.Rows[rowIndex];

            // Unhide the row and set its height
            row.IsHidden = false;
            row.Height = 20;

            // Save the workbook as a PDF document
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
