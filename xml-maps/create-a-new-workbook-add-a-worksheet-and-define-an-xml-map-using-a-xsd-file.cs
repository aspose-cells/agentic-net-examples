// Title: Create an Excel workbook, add a worksheet, and attach an XML map from an XSD schema using Aspose.Cells for .NET (C#)
// Common Searches: c# how to add an XML map from an XSD file to an Aspose.Cells workbook | aspnetcells create a new worksheet and bind an XSD schema as an XML map | export Excel worksheet data to XML using a defined XML map in Aspose.Cells | save a workbook with an embedded XML map using Aspose.Cells for .NET | sample code for adding an XML map and exporting data with Aspose.Cells C#
// Tags: aspnetcells create workbook xml map | c# add worksheet aspnetcells | aspnetcells bind xsd schema as xml map | aspnetcells export worksheet to xml | aspnetcells save workbook with xml map

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to instantiate a new Workbook with Aspose.Cells, add an additional Worksheet named "DataSheet", load an XSD schema file, create an XmlMap from that schema, populate sample cells, export the worksheet data to an XML file using the defined map, and finally save the workbook with the embedded XML map as an .xlsx file.
class Program
{
    static void Main()
    {
        // Path to the XSD (XML Schema) file that defines the XML map.
        // Ensure that "schema.xsd" exists in the same folder as the executable,
        // or provide a full absolute path.
        string xsdPath = "schema.xsd";

        if (!File.Exists(xsdPath))
        {
            Console.WriteLine($"Error: XSD file not found at \"{Path.GetFullPath(xsdPath)}\".");
            return;
        }

        try
        {
            // 1. Create a new workbook.
            Workbook workbook = new Workbook();

            // 2. Add a new worksheet (the workbook already contains one by default,
            //    but we add an extra sheet to demonstrate the operation).
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet worksheet = workbook.Worksheets[newSheetIndex];
            worksheet.Name = "DataSheet";

            // 3. Define an XML map using the XSD file.
            //    The XmlMaps collection belongs to the Worksheets collection.
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xsdPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

            Console.WriteLine("XML map successfully added.");
            Console.WriteLine($"Map Name          : {xmlMap.Name}");
            Console.WriteLine($"Root Element Name : {xmlMap.RootElementName}");

            // (Optional) Demonstrate that the map can be used to export data.
            // Add a few sample rows that match the schema (if any).
            worksheet.Cells["A1"].PutValue("SampleHeader1");
            worksheet.Cells["B1"].PutValue("SampleHeader2");
            worksheet.Cells["A2"].PutValue("Value1");
            worksheet.Cells["B2"].PutValue(123);

            // Export the worksheet data to an XML file using the defined map.
            string outputXmlPath = "ExportedData.xml";
            workbook.ExportXml(xmlMap.Name, outputXmlPath);
            Console.WriteLine($"Data exported to XML file: {Path.GetFullPath(outputXmlPath)}");

            // Save the workbook so you can inspect the embedded XML map.
            string outputWorkbookPath = "WorkbookWithXmlMap.xlsx";
            workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved as: {Path.GetFullPath(outputWorkbookPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An exception occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
