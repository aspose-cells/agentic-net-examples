using System;
using Aspose.Cells;

namespace AsposeCellsHtmlImportDemo
{
    class Program
    {
        static void Main()
        {
            // Create HTML load options; SupportDivTag defaults to false, so DIV layout will not be preserved
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML file into a workbook using the specified options
            Workbook workbook = new Workbook("input.html", loadOptions);

            // (Optional) Access the first worksheet to verify data
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("First cell value: " + sheet.Cells["A1"].StringValue);

            // Save the workbook to an Excel file
            workbook.Save("output.xlsx");
        }
    }
}