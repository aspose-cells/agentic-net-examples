using System;
using Aspose.Cells;

namespace AsposeCellsTxtToXmlDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Load the TXT file ----------
            // Create load options with a custom delimiter (e.g., '|')
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.Separator = '|';               // custom delimiter character
            // Load the TXT file into a workbook
            Workbook workbook = new Workbook("input.txt", loadOptions);

            // ---------- Split the text in the first column ----------
            Worksheet sheet = workbook.Worksheets[0];
            // Determine how many rows contain data in column A
            int totalRows = sheet.Cells.MaxDataRow + 1; // zero‑based index + 1
            // Create split options (same delimiter as used for loading)
            TxtLoadOptions splitOptions = new TxtLoadOptions();
            splitOptions.Separator = '|';
            // Perform TextToColumns on column A (row 0, column 0)
            sheet.Cells.TextToColumns(0, 0, totalRows, splitOptions);

            // ---------- Export the workbook data to XML ----------
            // The workbook must contain an XML map. For this demo we assume a map named "Map1" exists.
            // Export the XML data linked to that map.
            workbook.ExportXml("Map1", "output.xml");

            Console.WriteLine("TXT file processed, columns split, and XML exported to 'output.xml'.");
        }
    }
}