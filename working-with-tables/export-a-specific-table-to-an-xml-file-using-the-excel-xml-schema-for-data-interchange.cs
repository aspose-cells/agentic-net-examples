// Title: Export a Table Range to Excel XML with Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, populate a product table, configure XmlSaveOptions.ExportArea for A1:B4, optionally set the sheet name as the XML element, and save the selected range as an Excel‑compatible XML file (TableExport.xml) using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# XmlSaveOptions | Export range to XML | Excel XML schema | Export table to XML | Save worksheet as XML | CellArea ExportArea | Aspose.Cells example
// Common Searches: Aspose.Cells export specific range to XML C# | How to use XmlSaveOptions ExportArea in .NET | Save Excel table as XML file using Aspose.Cells | C# code to export A1:B4 to Excel XML | Excel XML schema export with Aspose.Cells
// Developer Intent: Save only the defined table (A1:B4) as an Excel XML file.
// Use Cases: Create an XML data feed for a product catalog by exporting a subset of worksheet data. | Generate lightweight XML reports that contain only pricing information from a larger spreadsheet. | Provide an Excel‑compatible XML interchange format for downstream systems that consume specific table ranges.
// AI Prompts: Generate code to export multiple tables from a workbook to separate XML files with Aspose.Cells. | Explain how to add custom XML namespaces when saving a worksheet using XmlSaveOptions. | Show how to read the produced TableExport.xml and validate it against the Excel XML schema.

using System;
using Aspose.Cells;

// Shows how to build a workbook, populate a product table, configure XmlSaveOptions.ExportArea for A1:B4, optionally set the sheet name as the XML element, and save the selected range as an Excel‑compatible XML file (TableExport.xml) using Aspose.Cells for .NET.
class ExportTableToXml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with a sample table (header + data)
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Price");
        worksheet.Cells["A2"].PutValue("Laptop");
        worksheet.Cells["B2"].PutValue(999.99);
        worksheet.Cells["A3"].PutValue("Phone");
        worksheet.Cells["B3"].PutValue(699.99);
        worksheet.Cells["A4"].PutValue("Tablet");
        worksheet.Cells["B4"].PutValue(450.75);

        // Configure XML save options to export only the defined table area
        XmlSaveOptions xmlOptions = new XmlSaveOptions
        {
            // Define the range A1:B4 (rows 0‑3, columns 0‑1)
            ExportArea = new CellArea { StartRow = 0, EndRow = 3, StartColumn = 0, EndColumn = 1 },
            // Export the sheet name as the XML element name (optional)
            SheetNameAsElementName = true
        };

        // Save the workbook as an XML file using the Excel XML schema
        workbook.Save("TableExport.xml", xmlOptions);

        Console.WriteLine("Table exported successfully to TableExport.xml");
    }
}
