// Title: How to load an existing .xlsx workbook and attach an XML schema map with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads a .xlsx file with Aspose.Cells, reads an XSD schema, creates an XML map, links it to the first worksheet, and saves the updated workbook. | Show how to use the Workbook.XmlMaps.Add method to bind an XSD file as an XML map in Aspose.Cells for .NET. | Provide a robust C# example with file existence checks and exception handling for adding an XML schema map to an Excel workbook.
// Common Searches: asp.net aspose.cells add xml map to existing excel workbook c# | c# load xsd schema and bind to worksheet using aspose.cells | how to associate an XML schema map with the first sheet in an .xlsx file via Aspose.Cells | save excel file with attached xml map using Aspose.Cells for .NET
// Tags: Workbook.XmlMaps.Add XSD schema | add XML map to Excel workbook using Aspose.Cells | bind XSD schema to worksheet Aspose.Cells | set XmlMapIndex on first worksheet Aspose.Cells | save workbook with attached XML map Aspose.Cells

using System;
using System.IO;
using System.Xml.Schema;
using Aspose.Cells;

// The example checks that both input.xlsx and schema.xsd exist, loads the workbook with Aspose.Cells, reads the XSD schema, adds it as an XML map named "MyMap", assigns the map to the first worksheet, and saves the modified workbook as output.xlsx, handling any errors that may occur.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string schemaPath = "schema.xsd";
        const string outputPath = "output.xlsx";

        // Verify required files exist to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input workbook not found: {inputPath}");
            return;
        }

        if (!File.Exists(schemaPath))
        {
            Console.WriteLine($"XML schema file not found: {schemaPath}");
            return;
        }

        try
        {
            // Load the workbook from an existing .xlsx file
            Workbook workbook = new Workbook(inputPath);

            // Load the XML schema from an .xsd file
            XmlSchema schema;
            using (FileStream fs = new FileStream(schemaPath, FileMode.Open, FileAccess.Read))
            {
                schema = XmlSchema.Read(fs, null);
            }

            // Add the schema as an XML map to the workbook (using dynamic to handle API variations)
            dynamic dynWorkbook = workbook;
            int mapIndex = dynWorkbook.XmlMaps.Add("MyMap", schema);

            // Associate the map with the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            dynamic dynSheet = sheet;
            dynSheet.XmlMapIndex = mapIndex;

            // Save the workbook with the attached XML schema map
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
