using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapImport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook(); // workbook-create

            // Add an XML map to the workbook using the XML schema file
            // Replace "schema.xsd" with the actual path to your XSD file
            int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd"); // xmlmapcollection-add

            // Access the newly added XmlMap (optional, e.g., to set a friendly name)
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex]; // worksheet-access (via collection index)

            // Optionally set a custom name for the map
            xmlMap.Name = "MyXmlMap";

            // Save the workbook to verify that the XML map has been added
            workbook.Save("WorkbookWithXmlMap.xlsx", SaveFormat.Xlsx); // workbook-save
        }
    }
}