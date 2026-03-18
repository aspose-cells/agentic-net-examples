using System;
using Aspose.Cells;

namespace AsposeCellsDivTagExample
{
    class Program
    {
        static void Main()
        {
            // Path to the HTML file that contains <div> tags
            string htmlPath = "input.html";

            // Enable support for <div> tag layout during HTML import
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
            loadOptions.SupportDivTag = true;

            // Load the HTML file into a workbook using the configured options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the workbook as an XLSX file, preserving the layout derived from <div> tags
            string xlsxPath = "output.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx);
        }
    }
}