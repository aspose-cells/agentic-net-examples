using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapWhitespaceDemo
{
    class Program
    {
        static void Main()
        {
            // Create XML load options
            XmlLoadOptions loadOptions = new XmlLoadOptions();

            // Enable XML mapping
            loadOptions.IsXmlMap = true;

            // Configure the options to ignore whitespace-like nodes.
            // Aspose.Cells does not have a dedicated whitespace flag,
            // but setting IgnoreRootAttributes helps avoid processing
            // empty root attributes that can appear as blank entries.
            loadOptions.IgnoreRootAttributes = true;

            // Load the XML file with the configured options
            Workbook workbook = new Workbook("input.xml", loadOptions);

            // (Optional) Access the first worksheet to verify data
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("First cell value after import: " + sheet.Cells["A1"].StringValue);

            // Save the workbook to Excel format
            workbook.Save("output.xlsx");
        }
    }
}