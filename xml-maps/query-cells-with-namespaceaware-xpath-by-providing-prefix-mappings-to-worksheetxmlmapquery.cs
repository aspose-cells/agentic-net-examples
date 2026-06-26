using System;
using System.Collections.Generic;
using Aspose.Cells;

class NamespaceAwareXmlMapQueryDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Sample XML containing a namespace with prefix 'ns'
            string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<ns:Root xmlns:ns='http://example.com/schema'>
    <ns:Data>
        <ns:Item>Value1</ns:Item>
        <ns:Item>Value2</ns:Item>
    </ns:Data>
</ns:Root>";

            // Import the XML into the first worksheet.
            // This automatically creates an XML map for the imported data.
            workbook.ImportXml(xml, "Sheet1", 0, 0);

            // Get the first worksheet (where the XML was imported)
            Worksheet worksheet = workbook.Worksheets[0];

            // The imported XML data is placed starting at cell A1.
            // For this simple example we read the first column until an empty cell is found.
            Console.WriteLine("Imported XML values:");
            int row = 0;
            while (true)
            {
                string value = worksheet.Cells[row, 0].StringValue;
                if (string.IsNullOrEmpty(value))
                    break;

                Console.WriteLine($"Row {row + 1}: {value}");
                row++;
            }

            // Save the workbook (optional)
            string outputPath = "NamespaceAwareQuery.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}