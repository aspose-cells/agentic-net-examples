// Title: Embedding a Book Catalog XSD Schema as a Custom XML Part in an Excel Workbook with Aspose.Cells (C#)
// Common Searches: C# Aspose.Cells add custom XML part to Excel workbook using ContentTypeProperties | How to embed an XSD schema into an .xlsx file with Aspose.Cells in .NET | Add a custom XML schema named bookCatalog.xsd to a workbook programmatically | Using Workbook.ContentTypeProperties.Add to store XML schema in Excel via Aspose.Cells | Aspose.Cells example for adding custom XML part for book catalog
// Tags: Aspose.Cells ContentTypeProperties custom XML insertion | XSD schema embedding in Excel workbook C# | C# add XML part to .xlsx with Aspose | book catalog schema integration Aspose.Cells | Workbook.ContentTypeProperties Add usage example

using Aspose.Cells;
using System;

// The sample creates a new Workbook, defines a book catalog XSD schema, adds it as a custom XML part named "bookCatalog.xsd" via Workbook.ContentTypeProperties.Add, optionally creates a worksheet, and saves the file as BookCatalogWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define the book catalog XSD schema as a string
        string bookCatalogSchema = @"<?xml version=""1.0"" encoding=""utf-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""catalog"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""book"" maxOccurs=""unbounded"">
          <xs:complexType>
            <xs:sequence>
              <xs:element name=""author"" type=""xs:string""/>
              <xs:element name=""title"" type=""xs:string""/>
              <xs:element name=""genre"" type=""xs:string""/>
              <xs:element name=""price"" type=""xs:float""/>
              <xs:element name=""publish_date"" type=""xs:date""/>
              <xs:element name=""description"" type=""xs:string""/>
            </xs:sequence>
            <xs:attribute name=""id"" type=""xs:string"" use=""required""/>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

        // Add the custom XML part to the workbook using ContentTypeProperties
        // The first argument is a unique name for the part, the second is the XML content.
        workbook.ContentTypeProperties.Add("bookCatalog.xsd", bookCatalogSchema);

        // (Optional) Add a worksheet for demonstration purposes
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Catalog";

        // Save the workbook to a file
        workbook.Save("BookCatalogWorkbook.xlsx");
    }
}
