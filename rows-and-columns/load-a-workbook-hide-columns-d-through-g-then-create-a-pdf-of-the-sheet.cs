// Title: Hide columns D through G in an Excel worksheet and convert the sheet to PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide specific columns before exporting to PDF | How to hide columns D to G in an Excel file using Aspose.Cells | Aspose.Cells .NET hide multiple columns and save as PDF example | Convert Excel worksheet to PDF after hiding columns with Aspose.Cells C# | Aspose.Cells hide column range and generate PDF from workbook
// Tags: column IsHidden property Aspose.Cells | Aspose.Cells PDF conversion C# | manipulate column visibility Aspose.Cells | Excel to PDF conversion with hidden columns | Aspose.Cells workbook save as PDF

using System;
using System.IO;
using Aspose.Cells;

// The C# sample loads an existing Excel workbook, hides columns D through G on the first worksheet by setting each column's IsHidden flag, and then saves the modified workbook as a PDF file using Aspose.Cells, with basic file‑existence checking and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide columns D (index 3) through G (index 6) inclusive
            for (int col = 3; col <= 6; col++)
            {
                // Use Columns collection to access column properties
                sheet.Cells.Columns[col].IsHidden = true;
            }

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
