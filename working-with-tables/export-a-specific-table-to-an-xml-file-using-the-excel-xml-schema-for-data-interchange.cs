// Title: Export Excel Table to XML with Aspose.Cells C# (XmlSaveOptions)
// Description: Creates a workbook, fills cells A1:B4 with product data, sets XmlSaveOptions (ExportArea, SheetNameAsElementName, DataAsAttribute), and saves the selected range as ProductsTable.xml using the Excel XML schema.
// Keywords: Aspose.Cells XML export C# | XmlSaveOptions ExportArea | save worksheet range as XML | Excel to XML data interchange | C# export table to XML file | Aspose.Cells generate XML from cells
// Common Searches: Aspose.Cells export specific range to XML C# | How to save Excel table as XML using XmlSaveOptions | C# convert worksheet area to XML file | Export Excel data to XML without schema Aspose | XmlSaveOptions example for table export
// Developer Intent: Produce an XML document from a defined cell block in a worksheet, adhering to the native Excel XML format.
// Use Cases: Create an XML feed for a product catalog directly from an Excel sheet. | Exchange a portion of a report with another system via XML. | Generate lightweight XML snapshots of dashboard data without external mapping files.
// AI Prompts: Show how to output column headers as XML attributes instead of elements. | Demonstrate using XmlMapName to map the exported range to a custom schema. | Explain exporting multiple non‑contiguous ranges into separate XML files with Aspose.Cells.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Creates a workbook, fills cells A1:B4 with product data, sets XmlSaveOptions (ExportArea, SheetNameAsElementName, DataAsAttribute), and saves the selected range as ProductsTable.xml using the Excel XML schema.
class ExportTableToXml
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate a sample table (including header row)
            ws.Cells["A1"].PutValue("Product");
            ws.Cells["B1"].PutValue("Price");
            ws.Cells["A2"].PutValue("Laptop");
            ws.Cells["B2"].PutValue(999.99);
            ws.Cells["A3"].PutValue("Phone");
            ws.Cells["B3"].PutValue(699.99);
            ws.Cells["A4"].PutValue("Tablet");
            ws.Cells["B4"].PutValue(450.75);

            // Configure XML save options – export the defined area as XML
            XmlSaveOptions saveOptions = new XmlSaveOptions
            {
                ExportArea = new CellArea { StartRow = 0, EndRow = 3, StartColumn = 0, EndColumn = 1 },
                SheetNameAsElementName = true,
                DataAsAttribute = false
                // No XmlMapName is set because we are not using an external schema
            };

            // Save the selected table as an XML file
            string outputPath = "ProductsTable.xml";
            wb.Save(outputPath, saveOptions);

            Console.WriteLine($"Table exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
