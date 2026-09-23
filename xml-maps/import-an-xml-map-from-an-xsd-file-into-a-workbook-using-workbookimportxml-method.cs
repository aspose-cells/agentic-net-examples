// Title: Import an XSD schema as an XML map into an Aspose.Cells workbook with C#
// AI Prompts: Generate C# code that creates a new Workbook, validates the presence of a .xsd file, and calls Workbook.ImportXml with loadMapOnly set to true to add only the XML map. | Show how to import just the XML map from a schema file and then save the workbook as an .xlsx file using Aspose.Cells. | Provide a C# example that demonstrates checking for the XSD file, importing the map without data, and handling exceptions during the ImportXml operation.
// Common Searches: c# aspocells import xml map from xsd file without loading data | Workbook.ImportXml loadMapOnly parameter example in C# | how to add an XML schema to a new Excel workbook using Aspose.Cells | verify xsd file exists before calling ImportXml in C# | save workbook after importing XML map with Aspose.Cells
// Tags: import xml map from xsd Aspose.Cells | Workbook.ImportXml loadMapOnly flag | c# add xml schema to excel workbook | aspose.cells import xml map without data | check xsd file existence before ImportXml

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates creating a Workbook, ensuring the XSD schema file exists, importing only the XML map using Workbook.ImportXml with the loadMapOnly flag, and saving the workbook to confirm the map was added.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Path to the XSD file that defines the XML map
            string xsdPath = "schema.xsd";

            // Ensure the XSD file exists before attempting to import
            if (!File.Exists(xsdPath))
            {
                Console.WriteLine($"Error: XSD file not found at '{xsdPath}'.");
                return;
            }

            // Import the XML map from the XSD file.
            // The ImportXml method expects integer flags for loadDataOnly (0 = false) and loadMapOnly (1 = true).
            workbook.ImportXml(xsdPath, string.Empty, 0, 1);

            // (Optional) If you have an XML data file to import using the map, you can do:
            // string xmlPath = "data.xml";
            // if (File.Exists(xmlPath))
            // {
            //     // Load data (1) without loading the map definition (0)
            //     workbook.ImportXml(xmlPath, string.Empty, 1, 0);
            // }

            // Save the workbook to verify that the map has been added
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
