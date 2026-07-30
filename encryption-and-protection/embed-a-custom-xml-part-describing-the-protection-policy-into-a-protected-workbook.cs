// Title: Embed Custom XML Policy into a Password‑Protected Excel Workbook with Aspose.Cells for .NET (C#)
// Description: Shows how to apply structure protection with a password, create a custom XML policy and XSD, embed the XML part into the workbook, and save the protected file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Excel protection | custom XML part | XSD schema | structure password | embed XML in workbook | metadata policy | protected workbook | Aspose.Cells for .NET
// Common Searches: embed custom xml part Aspose.Cells | add protection policy xml to Excel file | protect workbook structure with password C# | store metadata in Excel using custom XML | read custom xml part from protected workbook
// Developer Intent: Add a custom XML part that defines a protection policy to a workbook secured with a structure password.
// Use Cases: Include compliance metadata (author, edit rights) directly inside an Excel template. | Distribute protected workbooks that can be programmatically inspected for policy enforcement. | Enable downstream .NET applications to retrieve and validate the embedded protection policy without unlocking the file.
// AI Prompts: Generate C# code with Aspose.Cells to embed a custom XML policy and its XSD into a password‑protected workbook. | Show how to read and validate the embedded protection‑policy XML from a protected Excel file using Aspose.Cells for .NET. | Explain how to update the custom XML part after the workbook is saved while keeping the existing structure protection.

using System;
using System.Text;
using Aspose.Cells;

// Shows how to apply structure protection with a password, create a custom XML policy and XSD, embed the XML part into the workbook, and save the protected file using Aspose.Cells for .NET.
class EmbedCustomXmlProtectionPolicy
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "StrongPassword!123");

        // Define the custom XML that describes the protection policy
        string xmlData = @"<ProtectionPolicy xmlns=""http://example.com/policy"">
                               <Author>John Doe</Author>
                               <AllowEdit>false</AllowEdit>
                           </ProtectionPolicy>";

        // Define a simple XML schema for the custom part (optional but recommended)
        string xmlSchema = @"<?xml version=""1.0"" encoding=""UTF-8""?>
                             <xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema""
                                        targetNamespace=""http://example.com/policy""
                                        xmlns=""http://example.com/policy""
                                        elementFormDefault=""qualified"">
                               <xs:element name=""ProtectionPolicy"">
                                 <xs:complexType>
                                   <xs:sequence>
                                     <xs:element name=""Author"" type=""xs:string""/>
                                     <xs:element name=""AllowEdit"" type=""xs:boolean""/>
                                   </xs:sequence>
                                 </xs:complexType>
                               </xs:element>
                             </xs:schema>";

        // Add the custom XML part to the workbook
        workbook.CustomXmlParts.Add(
            Encoding.UTF8.GetBytes(xmlData),
            Encoding.UTF8.GetBytes(xmlSchema));

        // Save the protected workbook with the embedded custom XML part
        workbook.Save("ProtectedWorkbook_WithPolicy.xlsx");
    }
}
