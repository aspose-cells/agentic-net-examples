using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlDebug
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Prepare sample XML data
            string xmlData = "<root><item>DebugValue</item></root>";
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlData);

            // Add the custom XML part to the workbook (no schema provided)
            int partIndex = wb.CustomXmlParts.Add(xmlBytes, null);

            // Retrieve the added custom XML part
            CustomXmlPart customPart = wb.CustomXmlParts[partIndex];

            // Convert the stored byte[] back to a string for debugging
            string xmlContent = Encoding.UTF8.GetString(customPart.Data);

            // Log the XML content
            Console.WriteLine("Custom XML Part Content:");
            Console.WriteLine(xmlContent);
        }
    }
}