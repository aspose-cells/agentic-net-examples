using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlApostropheDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Enable automatic QuotePrefix handling so that a leading apostrophe
            //    is treated as a style flag rather than part of the cell text.
            workbook.Settings.QuotePrefixToStyle = true;

            // 3. Put a value that starts with a single quote.
            //    In Excel the leading apostrophe is not displayed; it only forces the cell
            //    to treat the content as text.
            Cell cell = cells["A1"];
            cell.PutValue("'Aspose");

            // 4. Verify that the QuotePrefix style is set (optional, for debugging)
            bool isQuotePrefix = cell.GetStyle().QuotePrefix;
            Console.WriteLine("QuotePrefix flag set: " + isQuotePrefix); // should be True

            // 5. Save the workbook as HTML using HtmlSaveOptions.
            //    ExportFormula is left as default (true) – not relevant here.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Ensure that HTML tags inside cells are not parsed (not needed for this case)
            htmlOptions.ParseHtmlTagInCell = false;

            string htmlPath = "WorkbookWithApostrophe.html";
            workbook.Save(htmlPath, htmlOptions);

            // 6. Read the generated HTML file.
            string htmlContent = File.ReadAllText(htmlPath);

            // 7. Verify that the leading apostrophe is correctly represented in the HTML.
            //    Aspose.Cells renders the apostrophe as &#39; when QuotePrefix is true.
            bool containsApostropheEntity = htmlContent.Contains("&#39;Aspose");
            bool containsLiteralApostrophe = htmlContent.Contains("'Aspose");

            Console.WriteLine("HTML contains '&#39;Aspose': " + containsApostropheEntity);
            Console.WriteLine("HTML contains literal \"'Aspose\": " + containsLiteralApostrophe);

            // 8. Output result of verification.
            if (containsApostropheEntity || containsLiteralApostrophe)
                Console.WriteLine("Leading apostrophe rendered correctly in HTML.");
            else
                Console.WriteLine("Leading apostrophe NOT found in HTML output.");
        }
    }
}