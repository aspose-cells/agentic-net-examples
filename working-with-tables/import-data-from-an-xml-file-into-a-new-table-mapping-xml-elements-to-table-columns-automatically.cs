// Title: Import XML into Excel as a Table with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new Workbook, use Aspose.Cells' ImportXml method to load an XML file directly into Sheet1 starting at cell A1, automatically map XML elements to columns, and save the result as an XLSX file.
// Keywords: Aspose.Cells ImportXml C# | XML to Excel conversion | load XML data into worksheet | automatic column mapping XML | create Excel table from XML | Aspose.Cells .NET example
// Common Searches: how to import xml into excel using aspose.cells | aspose.cells ImportXml C# example | map xml elements to excel columns automatically | convert xml file to xlsx programmatically | c# load xml data into worksheet with aspose
// Developer Intent: Load an XML file into a new worksheet, let Aspose.Cells map elements to columns, and export the data as an Excel workbook.
// Use Cases: Transform XML reports into searchable Excel tables for business analysts. | Automate daily ingestion of XML feeds into Excel for KPI dashboards. | Migrate legacy XML configuration files into spreadsheet format for auditing.
// AI Prompts: Generate C# code that reads an XML file and imports it into an Excel worksheet as a formatted table using Aspose.Cells. | Show how to apply a built‑in table style after importing XML data with Aspose.Cells ImportXml in .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create a new Workbook, use Aspose.Cells' ImportXml method to load an XML file directly into Sheet1 starting at cell A1, automatically map XML elements to columns, and save the result as an XLSX file.
class ImportXmlDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Path to the source XML file
        string xmlFilePath = "data.xml";

        // Import the XML data into the first worksheet (Sheet1) starting at cell A1 (row 0, column 0)
        // This automatically maps XML elements to table columns.
        workbook.ImportXml(xmlFilePath, "Sheet1", 0, 0);

        // Save the workbook with the imported data
        workbook.Save("ImportedFromXml.xlsx");
    }
}
