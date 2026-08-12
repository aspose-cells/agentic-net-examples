// Title: Auto‑Refresh XML Map in Aspose.Cells Workbook When Source XML Changes (C#)
// Description: Demonstrates how to create a workbook, add an external XML map, enable workbook‑wide IsRefreshAllConnections, set RefreshOnLoad for each data connection, and save the file so Excel automatically updates the XML map each time the workbook is opened.
// Keywords: Aspose.Cells XML map auto refresh | RefreshOnLoad C# | IsRefreshAllConnections Aspose.Cells | external XML connection Excel | auto update XML data workbook | Aspose.Cells XML map example | C# Excel XML map refresh
// Common Searches: Aspose.Cells refresh XML map on load | Enable automatic XML map update in .NET workbook | Set RefreshOnLoad for external connections Aspose.Cells | How to auto‑refresh XML data in Excel using Aspose.Cells | C# code to refresh XML map when source file changes
// Developer Intent: Configure a workbook so the linked XML map refreshes automatically whenever the source XML file is modified.
// Use Cases: Distribute a template that always reflects the latest values from a shared XML file. | Generate daily reports that pull current XML data without manual intervention. | Maintain synchronized data in collaborative workbooks where the XML source is frequently updated.
// AI Prompts: Provide C# code to add multiple XML maps to a workbook and enable auto‑refresh for each using Aspose.Cells. | Show how to toggle the RefreshOnLoad flag for a specific external connection after it has been created. | Explain steps to verify that an XML map refreshes automatically when opening the saved workbook in Excel.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsXmlMapRefreshDemo
{
    // Demonstrates how to create a workbook, add an external XML map, enable workbook‑wide IsRefreshAllConnections, set RefreshOnLoad for each data connection, and save the file so Excel automatically updates the XML map each time the workbook is opened.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add an XML map that points to an external XML file on disk
            // The XML file path can be absolute or relative
            string xmlFilePath = "data.xml";
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlFilePath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "DataMap";

            // Ensure that all external connections (including the XML map) are refreshed when the file is opened
            workbook.Worksheets.IsRefreshAllConnections = true;

            // Additionally, set the RefreshOnLoad flag for each external connection explicitly
            foreach (ExternalConnection conn in workbook.DataConnections)
            {
                conn.RefreshOnLoad = true;
            }

            // Save the workbook; when opened in Excel, it will automatically refresh the XML map data
            workbook.Save("WorkbookWithAutoRefresh.xml");
        }
    }
}
