using System;
using Aspose.Cells;

namespace AsposeCellsDivTagDemo
{
    class Program
    {
        static void Main()
        {
            // Create HTML load options and enable support for <div> tag layout
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.SupportDivTag = true;

            // Load the HTML file into a workbook using the specified options
            Workbook workbook = new Workbook("input.html", loadOptions);

            // Save the loaded workbook as an Excel file (XLSX)
            workbook.Save("output.xlsx");

            Console.WriteLine("HTML file loaded with <div> support and saved as XLSX.");
        }
    }
}