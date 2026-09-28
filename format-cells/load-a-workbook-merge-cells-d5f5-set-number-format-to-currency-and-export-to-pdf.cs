// Title: Merge cells D5 to F5, apply a built‑in currency format, and save the workbook as PDF using Aspose.Cells for .NET
// AI Prompts: Merge the range D5:F5, set the built‑in currency number format (ID 164), and export the worksheet to a PDF file with Aspose.Cells in C#. | Create a style with currency formatting, apply it to merged cells D5:F5, then save the workbook as a PDF using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells .NET how to merge a range and apply currency format before PDF conversion | C# set built‑in number format 164 on merged cells and export to PDF with Aspose.Cells | Save Excel workbook as PDF after merging cells D5:F5 and formatting as currency using Aspose.Cells | Apply style flag to merged cells and convert to PDF in Aspose.Cells C# example
// Tags: merge cells D5:F5 Aspose.Cells C# | apply built‑in currency number format Aspose.Cells | export worksheet to PDF Aspose.Cells | style flag all attributes Aspose.Cells | create style with number format ID 164 Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, merges cells D5:F5, applies the built‑in currency number format (ID 164) to the merged range, and saves the result as output.pdf.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells D5:F5 (row index 4, columns 3‑5)
            sheet.Cells.Merge(4, 3, 1, 3);

            // Create a style with a built‑in currency number format
            Style style = workbook.CreateStyle();
            style.Number = 164; // "$#,##0.00"

            // Apply the style to the merged range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("D5:F5");
            StyleFlag flag = new StyleFlag();
            flag.All = true; // Apply all style attributes (including number format)
            mergedRange.ApplyStyle(style, flag);

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
