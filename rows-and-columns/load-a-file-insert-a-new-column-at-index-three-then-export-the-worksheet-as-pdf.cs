// Title: Insert a column at index 3 in an Excel worksheet with Aspose.Cells for .NET and export the sheet to PDF
// AI Prompts: Load an existing .xlsx file, insert a new column at zero‑based index 3 in the first worksheet using Aspose.Cells, save the workbook to a temporary file, and convert it to a PDF. | Using Aspose.Cells for .NET, add a column at position 3 in a workbook, persist the change to a temporary file, then generate a PDF from that file while cleaning up the temporary file.
// Common Searches: Aspose.Cells C# insert column at specific index before PDF conversion | How to add a column to an Excel file and then export to PDF using Aspose.Cells | C# example inserting a column in the first worksheet and saving as PDF | Convert modified Excel workbook to PDF with a temporary file in Aspose.Cells
// Tags: Aspose.Cells column insertion C# | Aspose.Cells worksheet to PDF conversion | C# temporary workbook for PDF export | Excel worksheet modification prior to PDF generation | zero‑based column index insertion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads input.xlsx, inserts a column at zero‑based index 3 in the first worksheet, saves the workbook to a temporary file, converts the temporary file to output.pdf using ConversionUtility, and deletes the temporary file.
class Program
{
    static void Main()
    {
        // Paths for the original file, a temporary modified file, and the final PDF.
        string inputFile = "input.xlsx";
        string tempFile = "temp_modified.xlsx";
        string pdfFile = "output.pdf";

        // Load the existing workbook (load rule).
        Workbook workbook = new Workbook(inputFile);

        // Insert a new column at index 3 (zero‑based) in the first worksheet (insert column rule).
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.InsertColumn(3);

        // Save the modified workbook to a temporary file (save rule).
        workbook.Save(tempFile);

        // Convert the temporary Excel file to PDF using the provided ConversionUtility rule.
        ConversionUtility.Convert(tempFile, pdfFile);

        // Clean up the temporary file.
        if (File.Exists(tempFile))
        {
            File.Delete(tempFile);
        }

        Console.WriteLine("Worksheet exported to PDF successfully.");
    }
}
