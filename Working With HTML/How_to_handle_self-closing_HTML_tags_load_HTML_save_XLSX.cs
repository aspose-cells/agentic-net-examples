using System;
using Aspose.Cells;

namespace AsposeCellsHtmlSelfClosingTagDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains self‑closing tags (e.g., <br/>, <img/>)
            string htmlPath = "input.html";

            // Create HtmlLoadOptions to control how the HTML is parsed.
            // SupportDivTag is enabled to correctly handle <div> elements if they appear.
            // DeleteRedundantSpaces removes extra spaces that may be introduced by <br/> tags.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.SupportDivTag = true;
            loadOptions.DeleteRedundantSpaces = true;

            // Load the HTML file into a Workbook using the specified options.
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // (Optional) Access a cell to verify that the content was loaded correctly.
            // For example, print the value of cell A1.
            Console.WriteLine("Cell A1 value after loading HTML: " + workbook.Worksheets[0].Cells["A1"].StringValue);

            // Save the workbook as an XLSX file.
            string xlsxPath = "output.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine("HTML successfully converted to XLSX: " + xlsxPath);
        }
    }
}