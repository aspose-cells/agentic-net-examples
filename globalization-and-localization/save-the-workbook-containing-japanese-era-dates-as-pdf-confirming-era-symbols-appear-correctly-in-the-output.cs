// Title: Save an Excel workbook with Japanese era (ggge) date format as a PDF using Aspose.Cells for .NET
// Common Searches: how to apply Japanese era custom number format in Aspose.Cells and export to PDF | C# Aspose.Cells save workbook as PDF with Japanese era dates | export Excel to PDF preserving ggge date format with Aspose.Cells | Aspose.Cells PDF conversion keeping Japanese era symbols | set locale-specific date format for Japanese era before PDF export in Aspose.Cells
// Tags: ggge era formatting Aspose.Cells | export workbook to PDF Aspose.Cells C# | Japanese locale date format PDF conversion | custom number format PDF rendering Aspose.Cells | Aspose.Cells PDF output with locale-specific dates

using System;
using Aspose.Cells;

// The example creates a Workbook, inserts a DateTime into cell A1, applies the Japanese era custom number format "[$-ja-JP]ggge\"年\"m\"月\"d\"日\"", and saves the file as JapaneseEra.pdf, ensuring the era symbols appear correctly in the generated PDF.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define a sample date
        DateTime sampleDate = new DateTime(2023, 5, 1);

        // Put the date into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(sampleDate);

        // Apply Japanese era custom number format
        // Format: era name (gg), era year (e), followed by Japanese characters for year, month, day
        Style eraStyle = cell.GetStyle();
        eraStyle.Custom = "[$-ja-JP]ggge\"年\"m\"月\"d\"日\"";
        cell.SetStyle(eraStyle);

        // Save the workbook as PDF; the era symbols will be rendered in the PDF output
        workbook.Save("JapaneseEra.pdf", SaveFormat.Pdf);
    }
}
