// Title: Add a Protection‑Policy Custom XML Part to a Structure‑Protected Workbook with Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, apply structure protection with a password, define a protection‑policy XML document and its XSD schema, embed the XML as a custom part, and save the .xlsx file using Aspose.Cells for C#.
// Keywords: Aspose.Cells custom XML part | C# embed XML in Excel | structure protection password | Excel workbook protection policy | XSD schema custom part | protected workbook Aspose.Cells | .NET Excel encryption | store metadata in Excel file
// Common Searches: how to embed custom xml in a protected excel workbook using aspose.cells | add xml schema to a structure‑protected workbook c# | store protection policy as custom xml part in .xlsx | aspnet aspose.cells embed xml in password protected workbook | retrieve custom xml part from a protected Excel file
// Developer Intent: Embed a custom XML part that records the workbook's protection policy into a structure‑protected Excel file.
// Use Cases: Create a new workbook, protect its structure with a password, and embed a protection‑policy XML document with an XSD schema. | Open an existing structure‑protected workbook, replace or update the embedded custom XML part, and re‑save the file. | Validate a protection‑policy XML string against its XSD before adding it as a custom part to ensure schema compliance.
// AI Prompts: Generate C# code using Aspose.Cells to add a custom XML part and its XSD schema to a workbook that is protected with a structure password. | Show how to read, modify, and re‑embed a protection‑policy custom XML part in an existing password‑protected Excel file with Aspose.Cells. | Provide an example that validates an XML string against an XSD schema before embedding it as a custom XML part in a protected workbook.

using System;
using System.Text;
using Aspose.Cells;

// Demonstrates how to create a new Workbook, apply structure protection with a password, define a protection‑policy XML document and its XSD schema, embed the XML as a custom part, and save the .xlsx file using Aspose.Cells for C#.
class EmbedCustomXmlInProtectedWorkbook
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "pwd123");

        // Define custom XML data that describes the protection policy
        string xmlData = "<ProtectionPolicy xmlns=\"http://example.com/protection\">" +
                         "<Policy>StructureProtected</Policy>" +
                         "<Password>pwd123</Password>" +
                         "</ProtectionPolicy>";

        // (Optional) Define a simple XML schema for the custom part
        string xmlSchema = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                           "<xs:schema xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" " +
                           "targetNamespace=\"http://example.com/protection\" " +
                           "xmlns=\"http://example.com/protection\" elementFormDefault=\"qualified\">" +
                           "<xs:element name=\"ProtectionPolicy\">" +
                           "<xs:complexType>" +
                           "<xs:sequence>" +
                           "<xs:element name=\"Policy\" type=\"xs:string\"/>" +
                           "<xs:element name=\"Password\" type=\"xs:string\"/>" +
                           "</xs:sequence>" +
                           "</xs:complexType>" +
                           "</xs:element>" +
                           "</xs:schema>";

        // Add the custom XML part to the workbook
        workbook.CustomXmlParts.Add(Encoding.UTF8.GetBytes(xmlData), Encoding.UTF8.GetBytes(xmlSchema));

        // Save the protected workbook with the embedded custom XML part
        workbook.Save("ProtectedWithCustomXml.xlsx");
    }
}
