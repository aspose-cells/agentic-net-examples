using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // XML content to be stored in the custom XML part
        string xmlContent = "<root><item>Sample Data</item></root>";

        // Convert the XML string to a UTF‑8 byte array
        byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlContent);

        // Add the custom XML part to the workbook (no associated schema)
        int partIndex = workbook.CustomXmlParts.Add(xmlBytes, null);

        // Retrieve the newly added CustomXmlPart instance
        CustomXmlPart customPart = workbook.CustomXmlParts[partIndex];

        // Optionally assign a unique ID to the part
        customPart.ID = Guid.NewGuid().ToString();

        // Save the workbook (demonstrates that the custom XML part is persisted)
        workbook.Save("CustomXmlPartDemo.xlsx");
    }
}