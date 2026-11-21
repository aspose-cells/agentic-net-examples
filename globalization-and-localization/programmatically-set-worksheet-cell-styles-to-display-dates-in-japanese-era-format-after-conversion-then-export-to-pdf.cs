// Title: Format worksheet cells with Japanese era dates and export to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply Japanese era date format to cells | how to use custom number format for Japanese era in Aspose.Cells workbook | export Excel to PDF with Japanese era dates using Aspose.Cells .NET | set locale-specific date style ja-JP in Aspose.Cells C# | convert DateTime to Japanese era format in spreadsheet and save as PDF
// Tags: japanese era custom number format Aspose.Cells | locale ja-JP date style C# | apply style flag numberformat Aspose.Cells | export workbook to PDF Aspose.Cells | date formatting globalization Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, writes DateTime values to column A, defines a custom number format using the Japanese locale ([$-ja-JP]ggge"年"m"月"d"日"), applies the style with a StyleFlag that targets number formatting, and saves the result as a PDF file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "JapaneseEraDates";

        // Sample dates to be displayed
        DateTime[] sampleDates = new DateTime[]
        {
            new DateTime(2023, 5, 1),
            new DateTime(2024, 1, 15),
            new DateTime(2025, 12, 31)
        };

        // Write dates into column A (cells A1, A2, A3)
        for (int i = 0; i < sampleDates.Length; i++)
        {
            sheet.Cells[i, 0].PutValue(sampleDates[i]);
        }

        // Create a style with Japanese era custom number format
        // Format: era name + year + "年" + month + "月" + day + "日"
        // Locale identifier for Japanese (ja-JP) is used to enable era symbols
        Style japaneseEraStyle = workbook.CreateStyle();
        japaneseEraStyle.Custom = "[$-ja-JP]ggge\"年\"m\"月\"d\"日\"";

        // Apply the style to the date cells
        StyleFlag styleFlag = new StyleFlag();
        styleFlag.NumberFormat = true; // Only apply number format part of the style
        for (int i = 0; i < sampleDates.Length; i++)
        {
            sheet.Cells[i, 0].SetStyle(japaneseEraStyle, styleFlag);
        }

        // Export the workbook to PDF
        workbook.Save("JapaneseEraDates.pdf", SaveFormat.Pdf);
    }
}
