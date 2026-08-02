// Title: Export Excel worksheet to XML with data types and formatting using Aspose.Cells C#
// Description: Demonstrates how to create a workbook, fill cells with strings, integers, dates and doubles, apply date and currency styles, and save a selected range as an XML file while preserving native data types, header rows, and cell formatting via XmlSaveOptions.
// Keywords: Aspose.Cells XML export C# | preserve cell data types XML | export Excel to XML with formatting | XmlSaveOptions example | C# export worksheet as XML | date and currency style XML | Excel to XML Aspose.Cells
// Common Searches: Aspose.Cells export worksheet to XML with formatting | C# XmlSaveOptions preserve data types | How to keep date format when saving Excel as XML | Export specific range to XML using Aspose.Cells | Save Excel workbook as XML with header row
// Developer Intent: Save a worksheet (or a defined range) as XML while retaining original data types and applied styles.
// Use Cases: Generating XML reports that require exact numeric and date representations for downstream systems. | Creating data‑exchange files where column headers become XML element names and values keep their native types. | Integrating Excel data with legacy XML‑based applications, preserving formatting such as date and currency formats.
// AI Prompts: Show how to modify XmlSaveOptions to export cell values as attributes while keeping formatting. | Provide C# code to export multiple worksheets to separate XML files, preserving each sheet's styles. | Explain how to customize custom number formats before exporting to XML with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsXmlExport
{
    // Demonstrates how to create a workbook, fill cells with strings, integers, dates and doubles, apply date and currency styles, and save a selected range as an XML file while preserving native data types, header rows, and cell formatting via XmlSaveOptions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Populate cells with different data types
            sheet.Cells["A1"].PutValue("ID");          // Header (string)
            sheet.Cells["B1"].PutValue("Date");        // Header (string)
            sheet.Cells["C1"].PutValue("Amount");      // Header (string)

            sheet.Cells["A2"].PutValue(1);                         // Integer
            sheet.Cells["B2"].PutValue(new DateTime(2023, 1, 15)); // DateTime
            sheet.Cells["C2"].PutValue(1234.56);                   // Double

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(new DateTime(2023, 2, 20));
            sheet.Cells["C3"].PutValue(7890.12);

            // Apply formatting to preserve in the exported XML
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // Built‑in date format
            sheet.Cells["B2"].SetStyle(dateStyle);
            sheet.Cells["B3"].SetStyle(dateStyle);

            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 164; // Built‑in currency format
            sheet.Cells["C2"].SetStyle(currencyStyle);
            sheet.Cells["C3"].SetStyle(currencyStyle);

            // Configure XML save options to keep data types and formatting
            XmlSaveOptions saveOptions = new XmlSaveOptions
            {
                ExportArea = new CellArea { StartRow = 0, EndRow = 2, StartColumn = 0, EndColumn = 2 },
                HasHeaderRow = true,
                SheetNameAsElementName = true,
                DataAsAttribute = false // Export data as element values, preserving types
            };

            // Save the workbook as an XML file
            workbook.Save("SampleData.xml", saveOptions);
        }
    }
}
