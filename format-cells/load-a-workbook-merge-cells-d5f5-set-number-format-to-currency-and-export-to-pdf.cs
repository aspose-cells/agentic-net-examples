// Title: C# Aspose.Cells example: Merge cells D5‑F5, set a currency number format, and export the workbook to PDF
// Common Searches: how to merge a range of cells and apply a currency format before converting to PDF with Aspose.Cells in C# | Aspose.Cells C# merge D5:F5 and set custom number format $#,##0.00 | export an Excel workbook to PDF after formatting merged cells using Aspose.Cells | C# code to apply a currency style to merged cells and save as PDF with Aspose.Cells
// Tags: merge cells Aspose.Cells C# | currency number format style Aspose.Cells | export workbook to PDF Aspose.Cells | apply style to merged range Aspose.Cells | custom number format $#,##0.00 Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads or creates an Excel workbook, merges cells D5 to F5 on the first worksheet, applies a custom currency number format to the merged cell, and saves the workbook as a PDF file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                // No password is required for the sample file.
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a default worksheet.
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells D5:F5 (zero‑based row 4, column 3, 1 row, 3 columns)
            sheet.Cells.Merge(4, 3, 1, 3);

            // Create a style with currency number format
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "$#,##0.00";

            // Apply the style to the first cell of the merged range
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            sheet.Cells[4, 3].SetStyle(currencyStyle, flag);

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook processed successfully. PDF saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
