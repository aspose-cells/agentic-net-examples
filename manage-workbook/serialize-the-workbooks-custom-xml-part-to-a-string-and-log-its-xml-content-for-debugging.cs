using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add a sample custom XML part
        string xmlData = "<root><item>Sample</item></root>";
        byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlData);
        int partIndex = wb.CustomXmlParts.Add(xmlBytes, null);
        CustomXmlPart customPart = wb.CustomXmlParts[partIndex];

        // Serialize the custom XML part to a string
        string xmlContent = Encoding.UTF8.GetString(customPart.Data);

        // Log the XML content for debugging
        Console.WriteLine("Custom XML Part Content:");
        Console.WriteLine(xmlContent);
    }
}