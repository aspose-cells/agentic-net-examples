// Title: Export a worksheet to XML with Aspose.Cells in C# while preserving dates, numbers and formatting
// AI Prompts: Write C# code that creates a workbook, applies date and currency styles to cells, and saves it as XML using Aspose.Cells XmlSaveOptions configured to retain data types and number formats. | Demonstrate how to set XmlSaveOptions so that the first row is treated as a header and the worksheet name becomes the XML element name during export. | Adapt the sample to export only selected worksheet indexes to XML, ensuring that all applied cell styles are kept in the output.
// Common Searches: Aspose.Cells C# export worksheet to XML preserving cell formatting | How to keep date and currency formats when saving Excel as XML with Aspose.Cells | XmlSaveOptions DataAsAttribute false keep data types Aspose.Cells | Export Excel sheet with header row as XML elements using Aspose.Cells .NET | Save specific worksheets to XML with Aspose.Cells C#
// Tags: export worksheet to XML with XmlSaveOptions | preserve cell data types in XML output Aspose.Cells | apply date and currency number formats before XML export | include header row as XML elements Aspose.Cells | selective worksheet export to XML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds date and currency values with appropriate number formats, configures XmlSaveOptions to output data as elements, use the sheet name as the XML element, treat the first row as a header, and saves the workbook to an XML file while preserving cell data types and formatting.
class ExportWorksheetToXml
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Populate cells with different data types
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue(DateTime.Now);   // DateTime value
            sheet.Cells["B2"].PutValue(1234.56);        // Numeric value

            // Apply formatting to preserve data types in the XML output
            // Date format
            Style dateStyle = sheet.Cells["A2"].GetStyle();
            dateStyle.Number = 14; // Short date format
            sheet.Cells["A2"].SetStyle(dateStyle);

            // Currency format
            Style currencyStyle = sheet.Cells["B2"].GetStyle();
            currencyStyle.Number = 10; // Currency format
            sheet.Cells["B2"].SetStyle(currencyStyle);

            // Configure XML save options to retain formatting and data types
            XmlSaveOptions saveOptions = new XmlSaveOptions
            {
                DataAsAttribute = false,          // Export data as elements (preserves types)
                SheetNameAsElementName = true,    // Use sheet name as XML element name
                HasHeaderRow = true,              // First row is a header
                SheetIndexes = null               // Export all sheets
            };

            // Define output file path
            string outputPath = "ExportedData.xml";

            // Save the workbook as an XML file using the configured options
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully exported to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
