using System;
using Aspose.Cells;

namespace AsposeCellsHtmlSelfClosingTagDemo
{
    class Program
    {
        static void Main()
        {
            // Load the HTML file that contains self‑closing tags
            Workbook workbook = new Workbook("input.html");

            // Save the workbook to XLSX format
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("HTML with self‑closing tags imported and saved as XLSX successfully.");
        }
    }
}