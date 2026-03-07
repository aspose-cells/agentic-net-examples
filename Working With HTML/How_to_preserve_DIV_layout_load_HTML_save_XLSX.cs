using System;
using Aspose.Cells;

namespace PreserveDivLayoutExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains <div> tags
            string htmlPath = "input.html";

            // Create HTML load options and enable support for <div> layout
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
            loadOptions.SupportDivTag = true; // Preserve the visual layout defined by <div> elements

            // Load the HTML file into a workbook using the specified options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Optional: verify that data was loaded correctly
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("First cell value after load: " + sheet.Cells["A1"].StringValue);

            // Save the workbook as an XLSX file
            string xlsxPath = "output.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file with <div> layout saved as Excel: {xlsxPath}");
        }
    }
}