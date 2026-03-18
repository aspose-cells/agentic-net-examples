using System;
using Aspose.Cells;
using CellsRange = Aspose.Cells.Range;

namespace AsposeCellsQuotePrefixDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable automatic QuotePrefix style for strings that start with a single quote
            workbook.Settings.QuotePrefixToStyle = true;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Cells with leading single‑quote will have QuotePrefix style applied automatically
            sheet.Cells["A1"].PutValue("'TextWithQuote");
            sheet.Cells["A2"].PutValue("'12345"); // Looks like a number but should be treated as text

            // For a range, manually set the QuotePrefix style if needed
            // Example: apply QuotePrefix to cells B1:B3 regardless of the input string
            Style quoteStyle = workbook.CreateStyle();
            quoteStyle.QuotePrefix = true;

            // Apply the style to the range
            CellsRange range = sheet.Cells.CreateRange("B1:B3");
            range.ApplyStyle(quoteStyle, new StyleFlag { QuotePrefix = true });

            // Put values into the range (they will retain the leading‑quote formatting)
            sheet.Cells["B1"].PutValue("Sample1");
            sheet.Cells["B2"].PutValue("Sample2");
            sheet.Cells["B3"].PutValue("Sample3");

            // Save the workbook to XLSX
            string filePath = "QuotePrefixDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Load the workbook to verify that the QuotePrefix flag is preserved
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Output verification results
            Console.WriteLine("A1 QuotePrefix: " + loadedSheet.Cells["A1"].GetStyle().QuotePrefix);
            Console.WriteLine("A2 QuotePrefix: " + loadedSheet.Cells["A2"].GetStyle().QuotePrefix);
            Console.WriteLine("B1 QuotePrefix: " + loadedSheet.Cells["B1"].GetStyle().QuotePrefix);
            Console.WriteLine("B2 QuotePrefix: " + loadedSheet.Cells["B2"].GetStyle().QuotePrefix);
            Console.WriteLine("B3 QuotePrefix: " + loadedSheet.Cells["B3"].GetStyle().QuotePrefix);
        }
    }
}