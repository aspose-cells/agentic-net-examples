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

            // Optionally ignore root attributes (helps reduce unwanted nodes)
            loadOptions.IgnoreRootAttributes = true;

            // Load the XML file with the specified options.
            // The XML map will be created automatically based on the XML structure.
            Workbook workbook = new Workbook("input.xml", loadOptions);

            // At this point the XML data is imported.
            // Whitespace-only nodes are ignored by the loader, preventing blank entries.

            // Save the workbook to an Excel file.
            workbook.Save("output.xlsx");
        }
    }
}