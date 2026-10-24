// Title: Merge cells D5:F5, apply built‑in currency format, and save workbook as PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge a range and set currency number format before exporting to PDF | How to apply built‑in currency style to merged cells in Aspose.Cells and generate a PDF | Aspose.Cells .NET export merged cells D5 to F5 with number format as PDF | Set number format 164 on a merged range and convert workbook to PDF with Aspose.Cells | Save an Excel workbook to PDF after merging cells and applying currency formatting in C#
// Tags: merge cells with Aspose.Cells C# | apply currency number format Aspose.Cells | export workbook to PDF Aspose.Cells | built‑in number format 164 Aspose.Cells | merged range PDF conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file or creates a new workbook, merges cells D5:F5 on the first worksheet, applies the built‑in currency number format (Number 164) to the merged cells, and then saves the workbook as a PDF document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException ex) when (ex.Message.Contains("password"))
                {
                    Console.WriteLine("The workbook is password protected and cannot be opened without a password.");
                    return;
                }
            }
            else
            {
                // Create a new workbook with a sample value in D5.
                workbook = new Workbook();
                Worksheet wsNew = workbook.Worksheets[0];
                wsNew.Cells["D5"].PutValue(1234.56);
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells D5:F5 (zero‑based indices: row 4, column 3).
            sheet.Cells.Merge(4, 3, 1, 3);

            // Create a style with built‑in currency format.
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 164; // $#,##0.00

            // Apply the style to the top‑left cell of the merged range.
            sheet.Cells["D5"].SetStyle(currencyStyle);

            // Save the workbook as PDF.
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
