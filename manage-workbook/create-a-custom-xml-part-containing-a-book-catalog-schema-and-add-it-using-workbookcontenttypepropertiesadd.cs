// Title: Add a custom XML part with a book catalog schema to an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a new Workbook, defines a book catalog XML string, adds it as a custom XML part named "bookCatalog.xml" via Workbook.ContentTypeProperties.Add, and saves the workbook as an .xlsx file with Aspose.Cells. | Demonstrate how to embed any custom XML schema (for example, an employee directory) into an Excel workbook by using Workbook.ContentTypeProperties.Add in Aspose.Cells for .NET.
// Common Searches: asp.net add custom xml part to xlsx using Aspose.Cells ContentTypeProperties | how to embed book catalog XML into an Excel workbook with Aspose.Cells C# | store XML schema inside an Excel file using Aspose.Cells .NET example | using Workbook.ContentTypeProperties.Add to attach custom XML to a spreadsheet | Aspose.Cells custom XML part for Excel workbook tutorial
// Tags: add custom xml part with Workbook.ContentTypeProperties.Add | embed book catalog schema in Excel workbook using Aspose.Cells | store xml data in .xlsx via Aspose.Cells .NET | custom xml part for Excel file Aspose.Cells | content type properties add example C#

using Aspose.Cells;
using System;

// Creates a new Workbook, defines a book catalog XML string, adds it as a custom XML part named "bookCatalog.xml" using Workbook.ContentTypeProperties.Add, and saves the workbook as "BookCatalogWorkbook.xlsx".
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // XML content representing a simple book catalog schema
        string bookCatalogXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<catalog xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""
         xsi:noNamespaceSchemaLocation=""bookCatalog.xsd"">
  <book id=""bk101"">
    <author>Gambardella, Matthew</author>
    <title>XML Developer's Guide</title>
    <genre>Computer</genre>
    <price>44.95</price>
    <publish_date>2000-10-01</publish_date>
    <description>An in-depth look at creating applications with XML.</description>
  </book>
  <book id=""bk102"">
    <author>Ralls, Kim</author>
    <title>Midnight Rain</title>
    <genre>Fantasy</genre>
    <price>5.95</price>
    <publish_date>2000-12-16</publish_date>
    <description>A former architect battles corporate zombies.</description>
  </book>
</catalog>";

        // Add the custom XML part to the workbook.
        // The first argument is the part name, the second is the XML string.
        workbook.ContentTypeProperties.Add("bookCatalog.xml", bookCatalogXml);

        // Save the workbook to a file.
        workbook.Save("BookCatalogWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
