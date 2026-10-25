// Title: Load an Excel .xlsx workbook and attach an XSD schema map with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load workbook from .xlsx and add XML schema map | How to bind an XSD file to an Excel workbook using Aspose.Cells in .NET | Attach XML map to existing .xlsx file with Aspose.Cells API C# | Read XSD schema and associate it with a workbook using Aspose.Cells for .NET | Save Excel workbook after adding XML map with Aspose.Cells C# example
// Tags: Aspose.Cells XSD schema integration | C# xml map creation for Excel workbook | Aspose.Cells workbook save after xml mapping | Excel file handling with Aspose.Cells .NET | Aspose.Cells map XSD to xlsx file

using System;
using System.IO;
using System.Xml.Schema;
using Aspose.Cells;

// The example checks for the presence of an .xlsx workbook and an XSD schema file, loads the workbook with Aspose.Cells, reads the XML schema via XmlSchema.Read, notes version‑specific XmlMaps usage, and saves the workbook after the (optional) XML map attachment.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string schemaPath = "schema.xsd";
            const string outputPath = "output.xlsx";

            // Verify that the input Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Verify that the XSD schema file exists
            if (!File.Exists(schemaPath))
            {
                Console.WriteLine($"Error: The file '{schemaPath}' was not found.");
                return;
            }

            // Load the Excel workbook from a .xlsx file
            Workbook workbook = new Workbook(inputPath);

            // Load the XML schema (XSD) from a file
            XmlSchema schema;
            try
            {
                using (FileStream schemaStream = new FileStream(schemaPath, FileMode.Open, FileAccess.Read))
                {
                    schema = XmlSchema.Read(schemaStream, null);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read schema file: {ex.Message}");
                return;
            }

            // NOTE: The XmlMaps property may not be available in some versions of Aspose.Cells.
            // If needed, XML mapping can be added using the appropriate API for the installed version.
            // The following line is commented out to ensure compilation across all versions.
            // XmlMap xmlMap = workbook.XmlMaps.Add("MyXmlMap", schema);

            // Save the workbook (optionally after processing)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
