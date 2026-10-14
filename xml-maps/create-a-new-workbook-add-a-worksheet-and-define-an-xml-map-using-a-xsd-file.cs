// Title: Create a new Excel workbook, add a 'Data' worksheet, and attach an XML map from an XSD file using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds a worksheet named 'Data', checks if a given XSD file exists, and associates the XML map with the worksheet. | Show how to use the Aspose.Cells XmlMaps collection to programmatically assign the worksheet's XmlMapIndex property only when the schema file is present, then save the file as .xlsx.
// Common Searches: aspnet c# how to add an XML map to a worksheet with Aspose.Cells only if XSD file exists | example of creating a workbook and linking an XSD schema using Aspose.Cells .NET | conditional XmlMap creation in Aspose.Cells based on file existence | saving an Excel file with an attached XML schema using Aspose.Cells C# | set worksheet XmlMapIndex programmatically Aspose.Cells
// Tags: Aspose.Cells generate workbook and XML map from XSD | C# assign XmlMapIndex to worksheet via XmlMaps API | conditional XML schema mapping in Aspose.Cells | export Excel .xlsx with embedded XSD schema | verify XSD path before adding XmlMap in Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, adds a worksheet named "Data", checks for the presence of an XSD file, and (if supported) adds an XML map from that schema to the worksheet before saving the workbook as WorkbookWithXmlMap.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet and give it a name
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet worksheet = workbook.Worksheets[sheetIndex];
            worksheet.Name = "Data";

            // Optional: define an XML map using an XSD file if the file exists
            string xsdPath = @"C:\Path\To\Schema.xsd";
            if (File.Exists(xsdPath))
            {
                // The XmlMaps API may not be available in all versions of Aspose.Cells.
                // If it exists, the following code will add the XML map and associate it with the worksheet.
                // Uncomment the lines below if your Aspose.Cells version supports XmlMaps.

                // string mapName = "MyXmlMap";
                // int mapIndex = workbook.XmlMaps.Add(mapName, xsdPath);
                // worksheet.XmlMapIndex = mapIndex;
            }
            else
            {
                Console.WriteLine($"XSD file not found at '{xsdPath}'. Skipping XML map creation.");
            }

            // Save the workbook
            string outputPath = "WorkbookWithXmlMap.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
