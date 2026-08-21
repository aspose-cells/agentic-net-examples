// Title: Add a Custom XML Part with Book Catalog Schema to an Aspose.Cells Workbook via ContentTypeProperties (C#)
// Description: Creates a new Workbook, converts a book‑catalog XML document and its XSD schema to UTF‑8 byte arrays, adds them as a custom XML part, assigns a GUID ID, stores the part index in ContentTypeProperties, saves the file, and validates the part and property counts after reloading.
// Keywords: Aspose.Cells custom XML part | Workbook.ContentTypeProperties | C# add XML schema to Excel | embed XSD in workbook | custom XML part GUID | Excel metadata custom XML | Aspose.Cells example
// Common Searches: Aspose.Cells add custom XML part with XSD | How to use ContentTypeProperties in Aspose.Cells | Store custom XML part index in workbook | Retrieve custom XML part by ID Aspose.Cells | C# embed XML schema in Excel file
// Developer Intent: Embed an XML document and its XSD schema as a custom XML part in a workbook and link it through ContentTypeProperties.
// Use Cases: Package a book catalog XML inside an Excel file for data exchange between systems. | Assign a unique identifier to a custom XML part for later updates or retrieval. | Expose workbook metadata that points to a specific custom XML part for automated processing.
// AI Prompts: Generate C# code using Aspose.Cells to add a custom XML part from XML and XSD strings, assign a GUID, and record the part index in ContentTypeProperties. | Show how to load the saved workbook and verify the counts of CustomXmlParts and ContentTypeProperties. | Provide a snippet that reads the stored part index from ContentTypeProperties, retrieves the corresponding custom XML part, and outputs its XML content.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace AsposeCellsCustomXmlDemo
{
    // Creates a new Workbook, converts a book‑catalog XML document and its XSD schema to UTF‑8 byte arrays, adds them as a custom XML part, assigns a GUID ID, stores the part index in ContentTypeProperties, saves the file, and validates the part and property counts after reloading.
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // XML data representing a simple book catalog
            string xmlData = @"<catalog xmlns=""http://example.com/bookcatalog"">
                                   <book>
                                       <title>Sample Book</title>
                                   </book>
                               </catalog>";

            // XML schema (XSD) for the book catalog
            string schemaData = @"<?xml version=""1.0"" encoding=""UTF-8""?>
                                 <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'
                                            targetNamespace='http://example.com/bookcatalog'
                                            xmlns='http://example.com/bookcatalog'
                                            elementFormDefault='qualified'>
                                     <xs:element name='catalog'>
                                         <xs:complexType>
                                             <xs:sequence>
                                                 <xs:element name='book' maxOccurs='unbounded'>
                                                     <xs:complexType>
                                                         <xs:sequence>
                                                             <xs:element name='title' type='xs:string'/>
                                                         </xs:sequence>
                                                     </xs:complexType>
                                                 </xs:element>
                                             </xs:sequence>
                                         </xs:complexType>
                                     </xs:element>
                                 </xs:schema>";

            // Convert XML and schema strings to UTF-8 byte arrays
            byte[] xmlBytes = Encoding.UTF8.GetBytes(xmlData);
            byte[] schemaBytes = Encoding.UTF8.GetBytes(schemaData);

            // Add the custom XML part (data + associated schema) to the workbook
            int customXmlPartIndex = workbook.CustomXmlParts.Add(xmlBytes, schemaBytes);

            // Optionally, set a custom ID for the part (useful for later retrieval)
            CustomXmlPart addedPart = workbook.CustomXmlParts[customXmlPartIndex];
            addedPart.ID = Guid.NewGuid().ToString();

            // Add a content type property that references the custom XML part
            // Here we store the part's index as a string; the type is set to "string"
            workbook.ContentTypeProperties.Add("BookCatalogPartIndex", customXmlPartIndex.ToString(), "string");

            // Save the workbook to a file
            string outputPath = "BookCatalogWorkbook.xlsx";
            workbook.Save(outputPath);

            // Demonstrate that the custom XML part was added successfully
            Workbook loadedWorkbook = new Workbook(outputPath);
            Console.WriteLine($"Custom XML parts count after reload: {loadedWorkbook.CustomXmlParts.Count}");
            Console.WriteLine($"Content type properties count: {loadedWorkbook.ContentTypeProperties.Count}");
        }
    }
}
