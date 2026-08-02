// Title: C# – Convert Aspose.Cells Workbook to CSV with Japanese Era Date Formatting
// Description: This example shows how to set a workbook's region to Japan, apply a Japanese era custom date format (yyyy年M月d日), save as XLSX, and use Aspose.Cells ConversionUtility to export the file to CSV while preserving the locale‑specific date strings.
// Keywords: Aspose.Cells CSV conversion | C# Japanese era date format | locale specific CSV Aspose | Workbook.Settings.Region Japan | ConversionUtility XLSX to CSV | Japanese calendar formatting .NET | Excel to CSV with era dates | Aspose.Cells date locale
// Common Searches: Aspose.Cells export to CSV with Japanese era dates | How to keep Japanese calendar format when converting Excel to CSV in C# | ConversionUtility respect workbook region Japan | Set workbook region for locale‑aware CSV output | C# sample for Japanese date format in CSV using Aspose
// Developer Intent: Create a CSV file from an Excel workbook while retaining Japanese era date representations.
// Use Cases: Generate CSV reports for Japanese accounting systems that require era‑based dates. | Batch‑process localized Excel files for data pipelines that consume CSV with Japanese date strings. | Provide CSV exports for web applications targeting users in Japan, preserving familiar calendar format.
// AI Prompts: Write C# code using Aspose.Cells to convert an XLSX workbook to CSV and keep dates in Japanese era format. | Explain the effect of Workbook.Settings.Region = CountryCode.Japan on CSV output with ConversionUtility. | Show how to apply a custom Japanese calendar format to a cell before CSV conversion in Aspose.Cells.

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvJapaneseDate
{
    // This example shows how to set a workbook's region to Japan, apply a Japanese era custom date format (yyyy年M月d日), save as XLSX, and use Aspose.Cells ConversionUtility to export the file to CSV while preserving the locale‑specific date strings.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Set the workbook region to Japan to enable Japanese calendar handling
            wb.Settings.Region = CountryCode.Japan;

            // Access the first worksheet and a cell
            Worksheet sheet = wb.Worksheets[0];
            Cell dateCell = sheet.Cells["A1"];

            // Put a sample date value (e.g., 2023-09-15)
            dateCell.PutValue(new DateTime(2023, 9, 15));

            // Apply a custom format that uses the Japanese calendar year/month/day
            Style style = dateCell.GetStyle();
            style.Custom = "[$-F800]yyyy年m月d日"; // Japanese era format
            dateCell.SetStyle(style);

            // Save the workbook to a temporary XLSX file (required for ConversionUtility)
            string tempXlsxPath = "tempWorkbook.xlsx";
            wb.Save(tempXlsxPath, SaveFormat.Xlsx);

            // Convert the XLSX workbook to CSV using ConversionUtility.
            // The conversion respects the workbook's regional settings.
            string csvPath = "output.csv";
            ConversionUtility.Convert(tempXlsxPath, csvPath);

            Console.WriteLine($"Workbook converted to CSV at: {csvPath}");
        }
    }
}
