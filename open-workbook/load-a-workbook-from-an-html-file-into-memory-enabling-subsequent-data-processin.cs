using System;
using Aspose.Cells;

namespace AsposeCellsHtmlLoadDemo
{
    class Program
    {
        static void Main()
        {
            // Create HTML load options (default constructor)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Example: enable support for <div> tags and auto‑fit columns/rows
            loadOptions.SupportDivTag = true;
            loadOptions.AutoFitColsAndRows = true;

            // Load the HTML file into a Workbook using the options
            // This uses the Workbook(string, LoadOptions) constructor as defined in the API
            Workbook workbook = new Workbook("input.html", loadOptions);

            // Access the first worksheet for further processing
            Worksheet sheet = workbook.Worksheets[0];

            // Example: read a cell value (A1) and display it
            Console.WriteLine("Value of A1: " + sheet.Cells["A1"].StringValue);

            // Additional data manipulation can be performed here
            // e.g., put a new value into B2
            sheet.Cells["B2"].PutValue("Processed");

            // (Optional) Save the workbook to verify changes – uses the Workbook.Save method
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}