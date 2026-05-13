using System;
using Aspose.Cells;

namespace AsposeCellsTxtToXmlDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source TXT file (each line contains values separated by a custom delimiter, e.g., '|')
            string txtFilePath = "input.txt";

            // Load the TXT file into a workbook without splitting (use a separator that does not exist in the file)
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                Separator = '\0' // No splitting on load
            };
            Workbook workbook = new Workbook(txtFilePath, loadOptions);

            // Access the first worksheet and its cells collection
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the custom delimiter used in the TXT file (e.g., '|')
            TxtLoadOptions splitOptions = new TxtLoadOptions
            {
                Separator = '|'
            };

            // Split the data in column A (index 0) into multiple columns.
            int totalRows = cells.MaxRow + 1; // MaxRow is zero‑based, so add 1 for count
            cells.TextToColumns(0, 0, totalRows, splitOptions);

            // Export the processed data to XML.
            string xmlOutputPath = "output.xml";
            workbook.Save(xmlOutputPath, SaveFormat.Xml);

            Console.WriteLine($"Data exported to XML file: {xmlOutputPath}");
        }
    }
}