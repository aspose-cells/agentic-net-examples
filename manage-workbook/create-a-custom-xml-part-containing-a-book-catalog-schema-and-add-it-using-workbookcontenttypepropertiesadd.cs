// Title: Insert a custom XML part with a book catalog schema into an Excel workbook using Workbook.ContentTypeProperties.Add (C# Aspose.Cells)
// Common Searches: how to embed a custom XML part in an .xlsx file with Aspose.Cells C# | using Workbook.ContentTypeProperties.Add to add XML schema to an Excel workbook | Aspose.Cells C# example adding book catalog XML as a custom part | add custom XML content to an Excel workbook programmatically with Aspose.Cells | store XML catalog data inside an .xlsx file using Aspose.Cells API
// Tags: custom XML part insertion Aspose.Cells | Workbook.ContentTypeProperties.Add usage | embed book catalog schema Excel C# | add XML content to .xlsx programmatically | Aspose.Cells XML part example

using Aspose.Cells;
using System;

// The sample creates a new Workbook, defines a book catalog XML string, adds it as a custom XML part identified by "bookCatalog" via wb.ContentTypeProperties.Add, and saves the file as BookCatalog.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // XML schema representing a simple book catalog
        string bookCatalogXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<catalog>
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

        // Add the custom XML part to the workbook using ContentTypeProperties
        // "bookCatalog" is the identifier for the custom part
        wb.ContentTypeProperties.Add("bookCatalog", bookCatalogXml);

        // Save the workbook to a file
        wb.Save("BookCatalog.xlsx");
    }
}
