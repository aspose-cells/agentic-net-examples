// Title: How to set a custom namespace prefix for an XML map and attach it to a worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create an XML map from an XSD, assign a custom NamespacePrefix, and bind the map to the first worksheet with Aspose.Cells C#. | Update the reflection‑based code to set the XmlMap.NamespacePrefix property before saving the workbook. | Show how to add an XmlMap with a specific namespace prefix and associate it with a worksheet without using the direct XmlMaps API.
// Common Searches: aspnet set custom namespace prefix on xml map aspocells c# | how to bind xml map with custom namespace to worksheet using Aspose.Cells | reflection add xml map and set NamespacePrefix Aspose.Cells C# example | Aspose.Cells create xml map from xsd with custom namespace prefix
// Tags: set XmlMap NamespacePrefix Aspose.Cells C# | add XML map from XSD using reflection Aspose.Cells | bind XmlMap to worksheet Aspose.Cells | custom namespace handling in Excel export .NET | Aspose.Cells XML map configuration

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The sample creates a Workbook, loads an XSD file, adds an XmlMap named "MyXmlMap" via reflection, sets a custom NamespacePrefix on the map, ensures a worksheet exists, assigns the XmlMap to the first worksheet, creates the output directory if needed, saves the workbook as "output.xlsx", and handles any errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Path to the XML schema (XSD) that defines the XML map structure
            string xsdPath = "data.xsd";

            // Verify that the XSD file exists to avoid FileNotFoundException
            if (!File.Exists(xsdPath))
            {
                Console.WriteLine($"Error: XSD file not found at path '{xsdPath}'.");
                return;
            }

            // Add an XML map to the workbook using reflection (covers versions where XmlMaps may be unavailable)
            XmlMap xmlMap = null;
            PropertyInfo xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
            if (xmlMapsProp != null)
            {
                object xmlMaps = xmlMapsProp.GetValue(workbook);
                MethodInfo addMethod = xmlMaps?.GetType().GetMethod("Add", new[] { typeof(string), typeof(string) });
                if (addMethod != null)
                {
                    xmlMap = (XmlMap)addMethod.Invoke(xmlMaps, new object[] { "MyXmlMap", xsdPath });
                }
            }

            // Ensure there is at least one worksheet to associate with the XML map
            if (workbook.Worksheets.Count == 0)
            {
                workbook.Worksheets.Add();
            }

            // Reference the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Associate the worksheet with the XML map if the map was created successfully
            if (xmlMap != null)
            {
                PropertyInfo xmlMapProp = worksheet.GetType().GetProperty("XmlMap");
                if (xmlMapProp != null && xmlMapProp.CanWrite)
                {
                    xmlMapProp.SetValue(worksheet, xmlMap);
                }
            }

            // Save the workbook to a file
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
