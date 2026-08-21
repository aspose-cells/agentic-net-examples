// Title: C# – Create a Workbook, Add a Worksheet, and Define an XML Map from an XSD using Aspose.Cells
// Description: Demonstrates how to instantiate a Workbook, add a new Worksheet named "DataSheet", load an XSD schema via XmlMapCollection.Add, assign a friendly name to the XmlMap, and save the file as "WorkbookWithXmlMap.xlsx" with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# XML map | add worksheet Aspose.Cells | XmlMapCollection Add example | load XSD schema Aspose.Cells | save workbook with XML map | Aspose.Cells .NET tutorial | XML map from XSD C#
// Common Searches: Aspose.Cells add XML map from XSD | C# create workbook and worksheet Aspose.Cells | how to use XmlMapCollection.Add in .NET | save workbook with XML map Aspose.Cells | example code for XML map in Aspose.Cells
// Developer Intent: Create a workbook, insert a worksheet, and attach an XML map defined by an XSD schema.
// Use Cases: Generate a template workbook that includes an XML map for seamless XML import/export. | Prepare spreadsheets with predefined schemas to enable automated data binding between XML files and cells. | Build reporting pipelines where XML data is directly mapped to worksheet ranges for downstream processing.
// AI Prompts: Show C# code that loads an XSD file and adds it as an XML map to an Aspose.Cells workbook. | Provide an example of adding multiple worksheets, each with its own XML map, using Aspose.Cells. | Explain how to retrieve the name and schema information of an XmlMap after it has been added.

using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    // Demonstrates how to instantiate a Workbook, add a new Worksheet named "DataSheet", load an XSD schema via XmlMapCollection.Add, assign a friendly name to the XmlMap, and save the file as "WorkbookWithXmlMap.xlsx" with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: Workbook constructor)
            Workbook workbook = new Workbook();

            // Add a new worksheet to the workbook
            // The default workbook already contains one worksheet (index 0),
            // but we add an additional one to demonstrate the operation.
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = "DataSheet";

            // Define the path to the XSD file that describes the XML schema.
            // In a real scenario, ensure that "schema.xsd" exists at this location.
            string xsdPath = "schema.xsd";

            // Add the XML map to the workbook using the XmlMaps collection (rule: XmlMapCollection.Add)
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xsdPath);

            // Optionally set a friendly name for the XML map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "MyXmlMap";

            // Save the workbook to a file (lifecycle rule: Workbook.Save)
            workbook.Save("WorkbookWithXmlMap.xlsx");

            Console.WriteLine("Workbook created, worksheet added, and XML map defined successfully.");
        }
    }
}
