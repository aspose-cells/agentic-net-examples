using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create XML load options.
        // Set IsXmlMap to true to enable XML mapping.
        // Set IgnoreRootAttributes to true to ignore whitespace nodes (prevents blank entries).
        XmlLoadOptions loadOptions = new XmlLoadOptions
        {
            IsXmlMap = true,
            IgnoreRootAttributes = true
        };

        // Load the XML file using the configured options.
        Workbook workbook = new Workbook("input.xml", loadOptions);

        // Example: read a cell value to verify import.
        string firstCell = workbook.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("First cell value: " + firstCell);

        // Save the workbook to an Excel file.
        workbook.Save("output.xlsx");
    }
}