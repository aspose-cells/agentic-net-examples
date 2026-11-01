// Title: Format worksheet cells with Japanese era dates and export the workbook to PDF using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set Japanese era (ggge) custom number format for date cells | How to apply ja-JP locale date style to a range and save as PDF in Aspose.Cells | Export Excel dates in Reiwa/Heisei era to PDF with Aspose.Cells .NET | Apply StyleFlag number format only to date cells before PDF conversion using Aspose.Cells
// Tags: ggge custom number format c# | styleflag numberformat only range | culture aware date cell style | pdf export after cell formatting | aspocells era date conversion

using System;
using Aspose.Cells;

// The example creates a workbook, writes several DateTime values, defines a style with a Japanese era custom number format using the ja-JP locale (ggge"年"m"月"d"日"), applies the style to the date range via a StyleFlag that targets only the number format, and then saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Dates";

            // Sample dates to demonstrate Japanese era formatting
            DateTime[] dates = new DateTime[]
            {
                new DateTime(2023, 5, 1),   // Reiwa 5
                new DateTime(1995, 12, 15),// Heisei 7
                new DateTime(1989, 1, 8),  // Showa 64 (last day)
                new DateTime(1989, 1, 9)   // Heisei 1 (first day)
            };

            // Write dates into column A (cells A1:A4)
            for (int i = 0; i < dates.Length; i++)
            {
                Cell cell = sheet.Cells[i, 0];
                cell.PutValue(dates[i]);
            }

            // Create a style with Japanese era custom number format
            Style eraStyle = workbook.CreateStyle();
            // Custom format: era name (gg), era year (e), followed by Japanese characters for year, month, day
            // Locale identifier for Japanese (Japan) is 0x0411 (ja-JP)
            eraStyle.Custom = "[$-ja-JP]ggge\"年\"m\"月\"d\"日\"";

            // Apply the style to the range containing the dates
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true; // Apply only the number format part of the style
            Aspose.Cells.Range dateRange = sheet.Cells.CreateRange(0, 0, dates.Length, 1);
            dateRange.ApplyStyle(eraStyle, flag);

            // Save the workbook as PDF
            workbook.Save("JapaneseEraDates.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
