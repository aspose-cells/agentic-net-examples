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
        // The file "data.xml" should exist in the same folder as the executable
        int mapIndex = workbook.Worksheets.XmlMaps.Add("data.xml");
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
        xmlMap.Name = "DataMap";

        // Link a cell (A1) to an element in the XML map.
        // Adjust the XPath to match the structure of your XML file.
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Item");

        // Configure the external connection associated with the XML map
        // so that it refreshes automatically when the workbook is opened.
        if (workbook.DataConnections.Count > 0)
        {
            ExternalConnection connection = workbook.DataConnections[0];
            connection.RefreshOnLoad = true;
        }

        // Additionally, tell Excel to refresh all connections on opening.
        workbook.Worksheets.IsRefreshAllConnections = true;

        // Save the workbook
        workbook.Save("RefreshXmlMap.xlsx");
    }
}