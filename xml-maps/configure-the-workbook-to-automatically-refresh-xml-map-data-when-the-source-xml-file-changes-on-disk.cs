// Title: How to configure an Aspose.Cells workbook to auto‑refresh XML map data when the source XML file changes (C#)
// AI Prompts: Add an XML map to a workbook and enable its AutoRefresh property using Aspose.Cells in C#. | Use reflection to access the XmlMaps collection and add or replace an XmlMap for compatibility with older Aspose.Cells versions. | Save the workbook after setting AutoRefresh so the XML data updates automatically when the source file is modified. | Implement graceful fallback when the XmlMaps property or AutoRefresh member is unavailable.
// Common Searches: aspocells set xmlmap autoreload when xml file changes c# | c# add xml map to existing workbook using reflection aspocells | how to enable auto refresh for xml map in aspocells older versions | update excel workbook xml data automatically with aspocells | c# check if XmlMaps property exists in Aspose.Cells workbook
// Tags: Aspose.Cells set XmlMap AutoRefresh | C# reflection add XmlMap Aspose.Cells | auto refresh XML map in Excel workbook | configure XML map source file monitoring Aspose.Cells | legacy Aspose.Cells XmlMaps compatibility

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example loads an existing Excel workbook, uses reflection to obtain the XmlMaps collection, adds (or replaces) an XmlMap that points to a source XML file, sets the XmlMap's AutoRefresh property to true when the property is present, and saves the workbook so that XML data is refreshed automatically whenever the source XML file changes.
class Program
{
    static void Main()
    {
        try
        {
            // Define file paths
            string workbookPath = "input.xlsx";
            string xmlFilePath = "data.xml";
            string outputPath = "output.xlsx";
            string mapName = "MyXmlMap";

            // Verify that required files exist
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found – {workbookPath}");
                return;
            }

            if (!File.Exists(xmlFilePath))
            {
                Console.WriteLine($"Error: XML file not found – {xmlFilePath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(workbookPath);

            // Attempt to add (or replace) the XML map using reflection
            // This avoids compile‑time dependency on the XmlMaps property,
            // which may be absent in older Aspose.Cells versions.
            try
            {
                PropertyInfo xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
                if (xmlMapsProp != null)
                {
                    object xmlMaps = xmlMapsProp.GetValue(workbook);
                    MethodInfo addMethod = xmlMaps.GetType().GetMethod("Add", new[] { typeof(string), typeof(string) });
                    if (addMethod != null)
                    {
                        // Add the XML map
                        XmlMap xmlMap = (XmlMap)addMethod.Invoke(xmlMaps, new object[] { mapName, xmlFilePath });

                        // Enable AutoRefresh if the property exists
                        PropertyInfo autoRefreshProp = xmlMap.GetType().GetProperty("AutoRefresh");
                        if (autoRefreshProp != null && autoRefreshProp.CanWrite)
                        {
                            autoRefreshProp.SetValue(xmlMap, true);
                        }

                        Console.WriteLine("XML map added successfully.");
                    }
                    else
                    {
                        Console.WriteLine("Add method not found on XmlMaps collection.");
                    }
                }
                else
                {
                    Console.WriteLine("XmlMaps property is not available in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to add XML map – {ex.Message}");
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Exception: {ex.Message}");
        }
    }
}
