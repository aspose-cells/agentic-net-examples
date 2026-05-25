using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class BatchQuotePrefixUpdate
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with values that start with a single quote
            sheet.Cells["A1"].PutValue("'Text with quote");
            sheet.Cells["B2"].PutValue("'12345");
            sheet.Cells["C3"].PutValue("'Another example");

            // Create a style and enable the QuotePrefix property
            Style quoteStyle = workbook.CreateStyle();
            quoteStyle.QuotePrefix = true; // Mark cells as having a leading quote

            // Create a StyleFlag and enable only the QuotePrefix flag
            StyleFlag flag = new StyleFlag();
            flag.QuotePrefix = true; // Apply only the QuotePrefix setting

            // Apply the style to all cells in the worksheet using the flag
            sheet.Cells.ApplyStyle(quoteStyle, flag);

            // Save the workbook to verify the changes
            workbook.Save("BatchQuotePrefixUpdated.xlsx", SaveFormat.Xlsx);
        }
    }
}