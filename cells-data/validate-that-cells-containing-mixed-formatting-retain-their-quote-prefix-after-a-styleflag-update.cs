using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Put a value that starts with a single quote (quote prefix)
            cell.PutValue("'MixedFormatting");

            // Create an initial style with QuotePrefix = true and make the text bold
            Style initialStyle = workbook.CreateStyle();
            initialStyle.QuotePrefix = true;          // Enable quote prefix
            initialStyle.Font.IsBold = true;          // Additional formatting

            // Apply the initial style to the cell
            cell.SetStyle(initialStyle);

            // Verify that QuotePrefix is set initially
            Console.WriteLine("Initial QuotePrefix: " + cell.GetStyle().QuotePrefix); // Expected: True

            // Create a new style that changes only the font color
            Style newStyle = workbook.CreateStyle();
            newStyle.Font.Color = System.Drawing.Color.Red;

            // Create a StyleFlag that applies only the FontBold property (for demonstration)
            // Here we set FontBold flag to true to modify font boldness, leaving QuotePrefix untouched
            StyleFlag flag = new StyleFlag();
            flag.FontBold = true;   // Only font bold flag is enabled
            // All other flags, including QuotePrefix, remain false

            // Apply the new style with the flag; this should not affect the QuotePrefix
            cell.SetStyle(newStyle, flag);

            // After applying the style flag, check that QuotePrefix is still true
            bool quotePrefixAfterUpdate = cell.GetStyle().QuotePrefix;
            Console.WriteLine("QuotePrefix after StyleFlag update: " + quotePrefixAfterUpdate); // Expected: True

            // Save the workbook (optional, just to verify the result in Excel)
            workbook.Save("QuotePrefixValidation.xlsx");
        }
    }
}