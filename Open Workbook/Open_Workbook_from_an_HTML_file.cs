using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the HTML file to be loaded
        string htmlFilePath = "input.html";

        // Create HTML load options and configure desired settings
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.SupportDivTag = true;          // Enable support for <div> tags
        loadOptions.AutoFitColsAndRows = true;    // Auto‑fit columns and rows after loading

        // Load the HTML file into a Workbook using the (string, LoadOptions) constructor
        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

        // Access the first worksheet in the loaded workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Example: display the value of cell B1 (if it exists)
        Console.WriteLine("Cell B1 value: " + worksheet.Cells["B1"].StringValue);
    }
}