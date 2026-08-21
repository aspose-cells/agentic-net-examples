// Title: Save a formatted workbook to XLSX with Aspose.Cells for .NET while preserving styles
// Description: This C# example creates a workbook, inserts product data, defines a bold header with a light‑gray background, applies the style to the A1:B1 range using a StyleFlag, and saves the file as "FormattedOutput.xlsx" in XLSX format, keeping all formatting intact and handling possible exceptions.
// Keywords: Aspose.Cells save XLSX | preserve cell formatting .NET | C# Aspose.Cells StyleFlag | export styled Excel workbook | save workbook with header style | Aspose.Cells SaveFormat.Xlsx example | GitHub Aspose.Cells formatted export | Excel formatting retention Aspose
// Common Searches: Aspose.Cells save workbook with formatting | C# export styled Excel file using Aspose | preserve bold header and background when saving XLSX | how to use StyleFlag in Aspose.Cells | save formatted worksheet to XLSX .NET
// Developer Intent: Export a workbook to XLSX while retaining all applied cell styles.
// Use Cases: Generate a price list with a highlighted header and deliver it as a ready‑to‑print XLSX file. | Create a styled report (bold headers, colored cells) and share it without losing visual formatting. | Automate the production of Excel templates that require specific styling before distribution.
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, applies a bold, light‑gray header style to a range, and saves it as an XLSX file preserving all formatting. | Show how to use Aspose.Cells StyleFlag to apply all style attributes to a cell range and then persist the workbook. | Provide a robust Aspose.Cells example that saves a formatted worksheet to XLSX and includes exception handling.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSaveExample
{
    // This C# example creates a workbook, inserts product data, defines a bold header with a light‑gray background, applies the style to the A1:B1 range using a StyleFlag, and saves the file as "FormattedOutput.xlsx" in XLSX format, keeping all formatting intact and handling possible exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Put some data into cells
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.25);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.80);

                // Apply simple formatting: bold header row and set background color
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
                headerStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the header range A1:B1
                AsposeRange headerRange = sheet.Cells.CreateRange("A1:B1");
                headerRange.ApplyStyle(headerStyle, new StyleFlag { All = true });

                // Save the workbook to XLSX format, preserving all formatting
                workbook.Save("FormattedOutput.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
