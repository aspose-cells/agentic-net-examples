// Title: C# Example: Save Aspose.Cells Workbook to XLSX While Preserving All Formatting
// Description: Demonstrates how to create a workbook, add data, apply a bold header with a light‑gray background, and save it as an XLSX file using Aspose.Cells SaveFormat.Xlsx so that every style, color, and pattern remains intact.
// Keywords: Aspose.Cells C# save XLSX | preserve cell formatting Aspose.Cells | SaveFormat.Xlsx example | export formatted Excel .NET | Aspose.Cells workbook.Save formatting | C# Excel style retention | Aspose.Cells sample code
// Common Searches: save Aspose.Cells workbook as XLSX with formatting | C# Aspose.Cells keep styles on save | export formatted Excel file using Aspose.Cells | Aspose.Cells SaveFormat.Xlsx preserve colors | how to retain cell styles when saving Excel in .NET
// Developer Intent: Export a workbook to XLSX without losing any applied cell styles.
// Use Cases: Generate a price list with a styled header and share it as a ready‑to‑print XLSX file. | Create a report template, apply custom fonts and background colors, then distribute the formatted spreadsheet to stakeholders. | Automate data exports where downstream applications require the original Excel styling to remain visible.
// AI Prompts: Show C# code that adds borders, number formats, and font colors to a range, then saves the workbook to XLSX preserving all styles. | Provide an Aspose.Cells snippet that loads an existing XLSX, updates cell formatting, and saves without losing the original design. | Explain how to programmatically verify that header boldness and background color survive after calling workbook.Save with SaveFormat.Xlsx.

using System;
using Aspose.Cells;

namespace AsposeCellsSaveExample
{
    // Demonstrates how to create a workbook, add data, apply a bold header with a light‑gray background, and save it as an XLSX file using Aspose.Cells SaveFormat.Xlsx so that every style, color, and pattern remains intact.
    class Program
    {
        static void Main()
        {
            // Initialize a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.20);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Apply simple formatting: make header row bold and set background color
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the header cells
            sheet.Cells["A1"].SetStyle(headerStyle);
            sheet.Cells["B1"].SetStyle(headerStyle);

            // Save the workbook to XLSX format, preserving all formatting
            workbook.Save("FormattedOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}
