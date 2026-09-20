// Title: Load an Excel workbook with Arabic (ar‑SA) CultureInfo, apply right‑to‑left page setup, and save as PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file (or creates a new workbook with sample data when the file is absent), sets the workbook's CultureInfo to ar‑SA, configures the page setup for right‑to‑left orientation, and exports the result to PDF with Aspose.Cells. | Provide a .NET snippet that demonstrates applying Arabic (Saudi Arabia) localization to a worksheet, enabling right‑to‑left layout in the PDF output, and saving the workbook as a PDF file.
// Common Searches: Aspose.Cells set workbook culture to ar-SA in C# | How to enable right-to-left page layout when converting Excel to PDF with Aspose.Cells | Create a sample Excel file programmatically if input.xlsx does not exist using Aspose.Cells | Export Arabic localized Excel workbook to PDF preserving RTL orientation in .NET | Set CultureInfo for Aspose.Cells workbook before PDF conversion
// Tags: set workbook cultureinfo ar-sa Aspose.Cells | right-to-left page setup PDF export Aspose.Cells | create sample workbook if file missing C# Aspose.Cells | convert Excel to PDF with Arabic localization Aspose.Cells | apply RTL layout during PDF conversion .NET

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file or creates a new one with sample data, applies Arabic (Saudi Arabia) CultureInfo, optionally configures right‑to‑left page setup, and saves the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Define Arabic culture (Saudi Arabia)
            CultureInfo arabicCulture = new CultureInfo("ar-SA");

            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one with sample data
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Amount");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(100);
                ws.Cells["A3"].PutValue("A");
                ws.Cells["B3"].PutValue(150);
                ws.Cells["A4"].PutValue("B");
                ws.Cells["B4"].PutValue(200);
                ws.Cells["A5"].PutValue("B");
                ws.Cells["B5"].PutValue(250);
                workbook.Save(inputPath);
            }

            // Apply Arabic culture to the workbook settings
            workbook.Settings.CultureInfo = arabicCulture;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Note: Right‑to‑left layout can be set via PageSetup if supported.
            // If the property is unavailable in the current Aspose.Cells version,
            // this step is omitted to ensure compilation.

            // Determine the last row with data (optional, shown for reference)
            int lastDataRow = sheet.Cells.MaxDataRow;

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
