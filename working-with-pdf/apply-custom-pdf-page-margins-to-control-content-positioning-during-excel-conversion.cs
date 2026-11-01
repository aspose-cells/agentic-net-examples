// Title: Set custom PDF page margins for Excel-to-PDF conversion using Aspose.Cells in C#
// AI Prompts: Write C# code that configures left, right, top, and bottom margins in centimeters on a worksheet before saving it as a PDF with Aspose.Cells. | Demonstrate how to apply 0.5‑inch side margins and 0.75‑inch top/bottom margins to an Excel workbook during PDF export using Aspose.Cells.
// Common Searches: how to define worksheet margins in centimeters for Aspose.Cells PDF export C# | Aspose.Cells C# set 0.5 inch left and right margins when converting Excel to PDF | convert Excel file to PDF with 0.75 inch top and bottom margins using Aspose.Cells | C# example for adjusting page setup margins before saving workbook as PDF with Aspose.Cells
// Tags: worksheet page margins Aspose.Cells | excel to pdf margin configuration C# | pdf export margin settings Aspose.Cells | centimeter margin values Aspose.Cells | page setup margins conversion

using Aspose.Cells;
using System;
using System.IO;

// The program loads an Excel workbook, sets the first worksheet's left/right margins to 1.27 cm (0.5 in) and top/bottom margins to 1.905 cm (0.75 in), then saves the workbook as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the Excel workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (you can repeat for other sheets if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Set custom PDF page margins.
            // Aspose.Cells expects margins in centimeters.
            // 0.5 inch = 1.27 cm, 0.75 inch = 1.905 cm
            sheet.PageSetup.LeftMargin   = 1.27;   // 0.5 inch
            sheet.PageSetup.RightMargin  = 1.27;   // 0.5 inch
            sheet.PageSetup.TopMargin    = 1.905;  // 0.75 inch
            sheet.PageSetup.BottomMargin = 1.905;  // 0.75 inch

            // Convert and save the workbook as a PDF file with the specified margins
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
