// Title: Format Excel Dates in Japanese Era with Aspose.Cells .NET and Export to PDF
// Description: Demonstrates how to set a workbook's region to Japan, apply a custom Japanese era (gengō) number format to a cell, and convert the workbook to PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells Japanese era format | C# Excel date localization Japan | custom number format gengō | set workbook region Japan Aspose.Cells | Excel to PDF conversion Aspose.Cells .NET | ConversionUtility PDF export
// Common Searches: Aspose.Cells display dates in Japanese era | C# set workbook region to Japan for date formatting | custom number format for Japanese era in Excel | convert Excel with Japanese regional settings to PDF | Aspose.Cells PDF export with localized dates
// Developer Intent: Apply a Japanese era date format to Excel cells and generate a PDF from the workbook using Aspose.Cells for .NET.
// Use Cases: Produce financial statements for Japanese clients where dates must appear in gengō style before publishing as PDF. | Automate invoice creation for Japanese customers with era‑based dates and deliver each invoice as a PDF file. | Convert legacy Excel templates that store serial dates into PDF while preserving correct Japanese era representation.
// AI Prompts: Show C# code that sets the workbook region to Japan, formats a cell with the Japanese era custom number format, and exports the workbook to PDF using Aspose.Cells. | Provide an Aspose.Cells example that formats a range of cells with Japanese era dates and saves the result as a PDF. | Explain how to build the custom number format string for Japanese era display and ensure the PDF output reflects this formatting.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Demonstrates how to set a workbook's region to Japan, apply a custom Japanese era (gengō) number format to a cell, and convert the workbook to PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set the workbook's regional settings to Japan (required for Japanese era formatting)
        wb.Settings.Region = CountryCode.Japan;

        // Get the first worksheet and target cell A1
        Worksheet ws = wb.Worksheets[0];
        Cell cell = ws.Cells[0, 0]; // A1

        // Insert a serial date value (e.g., 44089 corresponds to 2020-09-15)
        cell.PutValue(44089);

        // Apply a custom number format that uses the Japanese era (gengō) representation
        Style style = cell.GetStyle();
        style.Custom = "[$-F800]yyyy年m月d日"; // Japanese era format
        cell.SetStyle(style);

        // Save the workbook to a temporary Excel file
        string excelPath = "JapaneseEra.xlsx";
        wb.Save(excelPath, SaveFormat.Xlsx);

        // Convert the saved Excel file to PDF using the provided ConversionUtility rule
        string pdfPath = "JapaneseEra.pdf";
        ConversionUtility.Convert(excelPath, pdfPath);

        Console.WriteLine("PDF generated at: " + pdfPath);
    }
}
