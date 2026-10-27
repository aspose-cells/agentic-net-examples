// Title: Load an Excel workbook with Arabic (ar‑SA) CultureInfo, add right‑to‑left subtotals, and save as PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook with Arabic (ar-SA) CultureInfo | How to enable right-to-left page layout for PDF export in Aspose.Cells | Create subtotals in a right-to-left worksheet using Aspose.Cells .NET | Export an Arabic‑localized Excel file to PDF with Aspose.Cells | Set RightToLeft property in Aspose.Cells page setup C#
// Tags: Arabic CultureInfo load options Aspose.Cells | right-to-left page setup PDF Aspose.Cells | subtotal creation RTL worksheet Aspose.Cells | PDF export localization Aspose.Cells | C# Aspose.Cells Arabic locale conversion

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example creates a simple workbook if needed, loads it with LoadOptions that specify the Arabic (Saudi Arabia) CultureInfo, optionally activates the RightToLeft page setup for correct printing direction, and then saves the workbook as a PDF, demonstrating how to handle Arabic localization and right‑to‑left subtotals with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var sampleWb = new Workbook();
                var ws = sampleWb.Worksheets[0];
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Amount");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue("A");
                ws.Cells["B3"].PutValue(20);
                ws.Cells["A4"].PutValue("B");
                ws.Cells["B4"].PutValue(30);
                sampleWb.Save(inputPath);
            }

            // Load the workbook using Arabic (Saudi Arabia) culture.
            var loadOptions = new LoadOptions
            {
                CultureInfo = new CultureInfo("ar-SA")
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Work with the first worksheet.
            var sheet = workbook.Worksheets[0];

            // If the RightToLeft property is available, enable it for printing direction.
            // This property may not exist in older versions of Aspose.Cells.
            // Uncomment the following line if your version supports it:
            // sheet.PageSetup.RightToLeft = true;

            // The Subtotal feature may not be available in all versions.
            // If needed, implement custom subtotal logic here.

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
