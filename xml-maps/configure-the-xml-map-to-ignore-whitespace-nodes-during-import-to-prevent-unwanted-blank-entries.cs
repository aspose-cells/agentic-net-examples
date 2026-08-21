// Title: Aspose.Cells C# – Ignore Whitespace Nodes in XML Map Import
// Description: Shows how to configure XmlLoadOptions (IsXmlMap = true, IgnoreRootAttributes = true) so that whitespace or empty XML nodes are omitted when loading an XML file into a Workbook, then export the map and save the workbook as XLSX.
// Keywords: Aspose.Cells | C# | XML map | ignore whitespace nodes | XmlLoadOptions | IgnoreRootAttributes | skip blank XML elements | import XML to Excel | remove empty entries | XML to Excel mapping
// Common Searches: Aspose.Cells ignore whitespace nodes | C# load XML into workbook without blank rows | XmlLoadOptions IgnoreRootAttributes example | skip empty XML elements Aspose.Cells | export XML map after import C#
// Developer Intent: Configure an XML map in Aspose.Cells to skip whitespace or empty nodes during import, preventing unwanted blank rows in the resulting workbook.
// Use Cases: Import XML files that contain formatting whitespace while keeping the worksheet free of empty rows. | Validate the import by exporting the first XML map and confirming that only meaningful data is present. | Save the cleaned workbook as an XLSX file for downstream processing or reporting.
// AI Prompts: Explain how XmlLoadOptions.IgnoreRootAttributes works and why it removes whitespace nodes in Aspose.Cells. | Provide a complete C# snippet that loads an XML file, ignores blank nodes with Aspose.Cells, and saves the workbook as XLSX. | Suggest alternative techniques to prevent empty entries when using XML maps in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to configure XmlLoadOptions (IsXmlMap = true, IgnoreRootAttributes = true) so that whitespace or empty XML nodes are omitted when loading an XML file into a Workbook, then export the map and save the workbook as XLSX.
class Program
{
    static void Main()
    {
        // Create XML load options and enable XML mapping.
        // Setting IgnoreRootAttributes helps to skip unnecessary whitespace nodes during import.
        XmlLoadOptions loadOptions = new XmlLoadOptions();
        loadOptions.IsXmlMap = true;
        loadOptions.IgnoreRootAttributes = true; // ignore whitespace-like nodes

        // Load the XML file into a workbook using the configured options.
        Workbook workbook = new Workbook("input.xml", loadOptions);

        // Export the XML data using the first available XML map (if any) to verify the import.
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            string mapName = workbook.Worksheets.XmlMaps[0].Name;
            workbook.ExportXml(mapName, "exported.xml");
        }

        // Save the workbook in Excel format.
        workbook.Save("result.xlsx");
    }
}
