using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet and a target cell (e.g., A1)
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Prefix the value with a single quote (apostrophe)
            cell.PutValue("'12345");

            // Enable QuotePrefix style so Excel treats the leading apostrophe as a text indicator
            Style style = cell.GetStyle();
            style.QuotePrefix = true;
            cell.SetStyle(style);

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}