// Title: Add a Protection‑Policy XML Custom Part to a Password‑Protected Workbook (Aspose.Cells for .NET)
// Description: Creates a new Workbook, applies structure protection with a password, builds an XML document and its XSD schema that describe the protection policy (password, type, author, read‑only flag), inserts the XML and schema as a custom XML part, saves the file, and reloads it to confirm the part persists and the workbook remains protected.
// Keywords: Aspose.Cells | C# custom XML part | Excel workbook protection | structure protection password | embed XML schema | policy metadata | save protected workbook | load workbook verify XML part
// Common Searches: Aspose.Cells add custom XML part to protected workbook | embed protection policy XML in Excel file using .NET | how to store workbook protection settings as XML in Aspose.Cells | C# protect workbook structure and include policy metadata | custom XML parts with XSD schema in Aspose.Cells
// Developer Intent: Insert a policy‑defining XML custom part into a password‑protected Excel workbook using Aspose.Cells.
// Use Cases: Archive password, protection type, author, and read‑only recommendation for compliance audits. | Enable downstream applications to read protection settings without opening the workbook. | Validate the embedded policy against an XSD schema to ensure consistency. | Programmatically verify that protection remains after the workbook is saved and reloaded.
// AI Prompts: Generate C# code with Aspose.Cells that protects a workbook and adds a custom XML part containing a protection‑policy XSD. | Show how to retrieve and parse the embedded protection‑policy XML from a saved workbook using Aspose.Cells. | Explain how to update the password value inside the embedded policy XML without removing workbook protection.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsProtectionPolicyExample
{
    // Creates a new Workbook, applies structure protection with a password, builds an XML document and its XSD schema that describe the protection policy (password, type, author, read‑only flag), inserts the XML and schema as a custom XML part, saves the file, and reloads it to confirm the part persists and the workbook remains protected.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Protect the workbook (structure protection)
            // -------------------------------------------------
            string workbookPassword = "WorkbookPwd123";
            workbook.Protect(ProtectionType.Structure, workbookPassword);

            // -------------------------------------------------
            // Define a custom XML part that describes the protection policy
            // -------------------------------------------------
            // Example XML describing the policy (you can customize as needed)
            string policyXml = @"
                <ProtectionPolicy xmlns=""http://example.com/protection"">
                    <WorkbookProtection>
                        <Password>" + workbookPassword + @"</Password>
                        <ProtectionType>Structure</ProtectionType>
                        <Author>John Doe</Author>
                        <RecommendReadOnly>true</RecommendReadOnly>
                    </WorkbookProtection>
                </ProtectionPolicy>";

            // Simple schema for the custom XML part (optional but shown for completeness)
            string schemaXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
                <xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"" targetNamespace=""http://example.com/protection"" xmlns=""http://example.com/protection"" elementFormDefault=""qualified"">
                    <xs:element name=""ProtectionPolicy"">
                        <xs:complexType>
                            <xs:sequence>
                                <xs:element name=""WorkbookProtection"">
                                    <xs:complexType>
                                        <xs:sequence>
                                            <xs:element name=""Password"" type=""xs:string""/>
                                            <xs:element name=""ProtectionType"" type=""xs:string""/>
                                            <xs:element name=""Author"" type=""xs:string""/>
                                            <xs:element name=""RecommendReadOnly"" type=""xs:boolean""/>
                                        </xs:sequence>
                                    </xs:complexType>
                                </xs:element>
                            </xs:sequence>
                        </xs:complexType>
                    </xs:element>
                </xs:schema>";

            // Add the custom XML part to the workbook
            workbook.CustomXmlParts.Add(
                Encoding.UTF8.GetBytes(policyXml),
                Encoding.UTF8.GetBytes(schemaXml));

            // -------------------------------------------------
            // Save the protected workbook with the custom XML part
            // -------------------------------------------------
            string outputPath = "ProtectedWorkbookWithPolicy.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Optional: Load the workbook back to verify the custom XML part count
            Workbook loaded = new Workbook(outputPath);
            Console.WriteLine($"Custom XML parts count after reload: {loaded.CustomXmlParts.Count}");
            Console.WriteLine($"Workbook is protected: {loaded.Settings.IsProtected}");
        }
    }
}
