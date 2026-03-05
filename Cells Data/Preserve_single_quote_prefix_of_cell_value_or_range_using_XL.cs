using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Enable automatic conversion of leading single quote to QuotePrefix style
            workbook.Settings.QuotePrefixToStyle = true;

            // Example 1: Let the setting handle the QuotePrefix automatically
            Cell cellA1 = cells["A1"];
            cellA1.PutValue("'AutoPrefix");
            Console.WriteLine("A1 QuotePrefix (auto): " + cellA1.GetStyle().QuotePrefix);

            // Example 2: Manually set QuotePrefix using Style
            Cell cellB2 = cells["B2"];
            cellB2.PutValue("'ManualPrefix");
            Style styleB2 = cellB2.GetStyle();
            styleB2.QuotePrefix = true;
            cellB2.SetStyle(styleB2);
            Console.WriteLine("B2 QuotePrefix (manual): " + cellB2.GetStyle().QuotePrefix);

            // Example 3: Apply QuotePrefix to a range using StyleFlag
            Style rangeStyle = workbook.CreateStyle();
            rangeStyle.QuotePrefix = true;

            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true;

            // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
            Aspose.Cells.Range range = cells.CreateRange("C3:D4");
            range.PutValue("'RangePrefix", false, false);
            range.ApplyStyle(rangeStyle, flag);

            Console.WriteLine("C3 QuotePrefix (range): " + cells["C3"].GetStyle().QuotePrefix);

            // Save the workbook as XLSX
            string filePath = "QuotePrefixDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {filePath}");

            // Load the saved workbook to verify that QuotePrefix persisted
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Cell loadedA1 = loadedSheet.Cells["A1"];
            Cell loadedB2 = loadedSheet.Cells["B2"];
            Cell loadedC3 = loadedSheet.Cells["C3"];

            Console.WriteLine("After reload:");
            Console.WriteLine("A1 QuotePrefix: " + loadedA1.GetStyle().QuotePrefix);
            Console.WriteLine("B2 QuotePrefix: " + loadedB2.GetStyle().QuotePrefix);
            Console.WriteLine("C3 QuotePrefix: " + loadedC3.GetStyle().QuotePrefix);
        }
    }
}