using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create HTML load options and enable support for <div> layout
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.SupportDivTag = true; // PreserveDivLayout equivalent

            // Load the HTML file using the specified options
            Workbook workbook = new Workbook("input.html", loadOptions);

            // Save the loaded workbook as an XLSX file
            workbook.Save("output.xlsx");
        }
    }
}