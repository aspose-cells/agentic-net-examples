using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add an XML map that points to the source XML file on disk
        // Replace "data.xml" with the actual path to your XML file or XSD schema
        int mapIndex = workbook.Worksheets.XmlMaps.Add("data.xml");
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "DataMap";

        // (Optional) Link a cell to an element in the XML map so that the cell shows XML data
        // Adjust the XPath to match your XML structure
        workbook.Worksheets[0].Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Item");

        // Ensure that any external connections (including the one created for the XML map)
        // are refreshed automatically when the workbook is opened
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            connection.RefreshOnLoad = true;
        }

        // Additionally, tell Excel to refresh all connections on opening the file
        workbook.Worksheets.IsRefreshAllConnections = true;

        // Save the workbook
        workbook.Save("WorkbookWithAutoRefresh.xlsx");
    }
}