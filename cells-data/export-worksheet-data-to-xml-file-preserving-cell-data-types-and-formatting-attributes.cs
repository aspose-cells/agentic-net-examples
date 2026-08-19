// Title: Export Excel Worksheet to XML with Data Types and Formatting – Aspose.Cells C#
// Description: Demonstrates how to create a workbook, fill cells with date, numeric, boolean and text values, apply a bold blue header style, and use XmlSaveOptions to export a defined range to an XML file while preserving cell data types, header information, and sheet‑level formatting.
// Keywords: Aspose.Cells | C# | Export to XML | XmlSaveOptions | preserve data types | cell formatting | header row XML | specific range export | Excel to XML conversion
// Common Searches: Aspose.Cells export worksheet to XML C# | keep Excel data types when saving as XML | XmlSaveOptions header formatting example | export selected cells to XML with Aspose | save workbook as XML with sheet name element
// Developer Intent: Generate an XML file from an Excel worksheet that retains original data types, header row, and applied styles.
// Use Cases: Provide typed XML feeds for APIs that consume dates, numbers, and booleans directly from Excel data. | Create XML reports where styled headers convey column meaning for downstream processing. | Exchange a specific cell block with legacy systems that require explicit type metadata in XML.
// AI Prompts: Show how to modify XmlSaveOptions to export cell values as attributes while still preserving data types. | Provide code to read the generated XML back into an Aspose.Cells workbook and keep the original formatting. | Explain how to iterate over all worksheets and save each as a separate XML file using XmlSaveOptions.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to create a workbook, fill cells with date, numeric, boolean and text values, apply a bold blue header style, and use XmlSaveOptions to export a defined range to an XML file while preserving cell data types, header information, and sheet‑level formatting.
class ExportWorksheetToXml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "SampleData";

        // Populate cells with different data types
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue(DateTime.Now);
        sheet.Cells["B2"].PutValue(1234.56);
        sheet.Cells["A3"].PutValue(true);
        sheet.Cells["B3"].PutValue("Text");

        // Apply formatting to header cells (bold blue font)
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.Color = Color.Blue;
        headerStyle.Font.IsBold = true;
        sheet.Cells["A1"].SetStyle(headerStyle);
        sheet.Cells["B1"].SetStyle(headerStyle);

        // Configure XmlSaveOptions to preserve data types and formatting
        XmlSaveOptions saveOptions = new XmlSaveOptions
        {
            // Export data as elements (not attributes) to keep type information clear
            DataAsAttribute = false,
            // Use sheet name as the XML element name
            SheetNameAsElementName = true,
            // Indicate that the exported range contains a header row
            HasHeaderRow = true,
            // Define the exact area to export (rows 0-2, columns 0-1)
            ExportArea = new CellArea { StartRow = 0, EndRow = 2, StartColumn = 0, EndColumn = 1 }
        };

        // Save the workbook as an XML file using the configured options
        workbook.Save("ExportedData.xml", saveOptions);
    }
}
