// Title: Create a Japanese era date style for Excel cells and save the worksheet as a PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that defines a custom number format using "ggge\"年\"m\"月\"d\"日\"" to display dates in the Japanese era, applies the style to a specific cell, and saves the workbook as a PDF. | Show how to use Aspose.Cells Style and StyleFlag objects to apply only the number format for a date cell, then export the worksheet to a PDF file. | Provide a step‑by‑step example that creates a workbook, inserts the current DateTime into cell A1, formats it with the Reiwa era pattern, and writes the result to JapaneseEraDate.pdf.
// Common Searches: aspnet format excel cell with japanese era (ggge) and export to pdf | c# aspose.cells custom number format for reiwa era date | how to apply japanese era date style to a worksheet cell using asp.net | save excel workbook as pdf while preserving localized date formats in asp.net | asp.net set styleflag numberformat true for date cell in aspose.cells
// Tags: Japanese era custom number format Aspose.Cells | apply styleflag numberformat to date cell C# | export workbook to pdf with localized date format | ggge date pattern Aspose.Cells | create style for era name in Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, writes the current date into cell A1, defines a custom style with the Japanese era format "ggge\"年\"m\"月\"d\"日\"", applies the style using a StyleFlag that targets only the number format, and saves the workbook as JapaneseEraDate.pdf.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a date value in cell A1 (example: today)
        Cell dateCell = sheet.Cells["A1"];
        dateCell.PutValue(DateTime.Now);

        // Create a style for Japanese era date format
        Style eraStyle = workbook.CreateStyle();
        // Custom number format: era name (gg), era year (e), month (m), day (d)
        // Example output: "令和3年9月10日"
        eraStyle.Custom = "ggge\"年\"m\"月\"d\"日\"";

        // Apply the style to the cell
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true; // Apply only number format
        dateCell.SetStyle(eraStyle, flag);

        // Export the worksheet to PDF
        workbook.Save("JapaneseEraDate.pdf", SaveFormat.Pdf);
    }
}
