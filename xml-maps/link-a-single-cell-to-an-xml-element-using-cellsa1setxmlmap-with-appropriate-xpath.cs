// Title: Bind cell A1 to an XML node using Cells.LinkToXmlMap (XPath) – Aspose.Cells for .NET
// Description: Creates a workbook, writes a simple XML file, adds it as an XML map, then connects cell A1 to the '/root/Item' node via Cells.LinkToXmlMap, optionally sets a placeholder value, and saves the workbook.
// Keywords: Aspose.Cells XML map C# | Cells.LinkToXmlMap example | XPath cell binding Aspose.Cells | link Excel cell to XML node | .NET Excel XML mapping
// Common Searches: Aspose.Cells bind single cell to XML element C# | Cells.LinkToXmlMap XPath usage | Create XML map from file and link cell in Aspose.Cells | How to map Excel cell to XML node with Aspose.Cells
// Developer Intent: Connect a worksheet cell to a specific XML element using an XML map and XPath.
// Use Cases: Import a value from an XML document directly into a designated cell for reporting. | Design an Excel template where certain cells are pre‑linked to XML nodes for seamless data exchange. | Maintain synchronization between Excel cells and XML elements when the workbook is edited.
// AI Prompts: Show how to link multiple cells to different XML nodes with Aspose.Cells in C#. | Explain how to update the underlying XML when a linked cell value changes. | Provide error‑handling patterns for missing or malformed XML files when adding an XML map.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlLinkDemo
{
    // Creates a workbook, writes a simple XML file, adds it as an XML map, then connects cell A1 to the '/root/Item' node via Cells.LinkToXmlMap, optionally sets a placeholder value, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Minimal XML content to be used for the map
                string xmlContent = "<root><Item>SampleValue</Item></root>";
                string xmlFilePath = "sample.xml";

                // Ensure the XML file exists before adding it as a map
                if (!File.Exists(xmlFilePath))
                {
                    File.WriteAllText(xmlFilePath, xmlContent);
                }

                // Add an XML map from the XML file (Aspose.Cells expects a file path)
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlFilePath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "DemoMap"; // Assign a friendly name to the map

                // Get the first worksheet and its cells collection
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Link cell A1 (row 0, column 0) to the XML element using XPath
                cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/root/Item");

                // Optional placeholder value
                cells["A1"].PutValue("Linked to XML");

                // Save the workbook
                string outputPath = "LinkedCell.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
