// Title: How to load an Excel workbook from a MemoryStream, add an XML map via reflection, and save it back to a MemoryStream in C# using Aspose.Cells
// AI Prompts: Write C# code that reads an XLSX file into a MemoryStream, creates an XML map from an XSD string using Aspose.Cells (with reflection as a fallback), attaches the map to the first worksheet, and returns the updated workbook as another MemoryStream. | Demonstrate step‑by‑step how to use reflection to invoke the XmlMaps.Add method when the XmlMaps API is not directly exposed, bind the resulting map to a worksheet, and then save the workbook to a byte array.
// Common Searches: asp.net load excel from byte array add xml map aspose.cells | c# add xml map to workbook using reflection aspose cells | save modified workbook to memory stream aspose.cells .net | bind xsd schema to worksheet with aspose.cells | fallback xmlmaps api not available aspose cells version
// Tags: excel workbook memory stream loading Aspose.Cells | xml map addition via reflection Aspose.Cells | xsd schema binding to worksheet Aspose.Cells | save workbook to memory stream C# | xmlmaps fallback handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for an input.xlsx file, creates it if missing, loads the workbook from a MemoryStream, uses reflection to add an XML map defined by an XSD string to the first worksheet, and then saves the modified workbook back into another MemoryStream before writing it to output.xlsx.
class XmlMapExample
{
    static void Main()
    {
        try
        {
            // Verify that the input file exists before attempting to read it
            const string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                // Create a new workbook if the source file does not exist
                var newWb = new Workbook();
                newWb.Save(inputPath, SaveFormat.Xlsx);
            }

            // Load workbook bytes
            byte[] inputBytes = File.ReadAllBytes(inputPath);

            // Load workbook from a memory stream
            using (MemoryStream inputStream = new MemoryStream(inputBytes))
            {
                Workbook workbook = new Workbook(inputStream);

                // Attempt to add an XML map using reflection (avoids compile‑time dependency on XmlMaps API)
                try
                {
                    // XML schema (XSD) as a string – replace with your actual schema definition
                    string xmlSchema = @"<?xml version=""1.0"" encoding=""utf-8""?>
<xs:schema xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:element name=""Root"">
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""Item"" type=""xs:string"" minOccurs=""0"" maxOccurs=""unbounded""/>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

                    // Check if the XmlMaps property exists
                    var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
                    if (xmlMapsProp != null)
                    {
                        var xmlMaps = xmlMapsProp.GetValue(workbook);
                        var addMethod = xmlMaps.GetType().GetMethod("Add", new[] { typeof(string), typeof(string) });
                        if (addMethod != null)
                        {
                            // Add the XML map and obtain its index
                            var mapIndexObj = addMethod.Invoke(xmlMaps, new object[] { "MyXmlMap", xmlSchema });
                            int mapIndex = Convert.ToInt32(mapIndexObj);

                            // Retrieve the first worksheet
                            var sheet = workbook.Worksheets[0];
                            var xmlMapProp = sheet.GetType().GetProperty("XmlMap");
                            if (xmlMapProp != null)
                            {
                                // Get the XmlMap instance from the collection via its indexer
                                var indexer = xmlMaps.GetType().GetProperty("Item", new[] { typeof(int) });
                                var xmlMap = indexer?.GetValue(xmlMaps, new object[] { mapIndex });

                                // Associate the map with the worksheet
                                xmlMapProp.SetValue(sheet, xmlMap);
                                Console.WriteLine("XML map added successfully.");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("XmlMaps API is not available in the current Aspose.Cells version.");
                    }
                }
                catch (Exception ex)
                {
                    // Log any issues related to XML mapping
                    Console.WriteLine($"Error while adding XML map: {ex.Message}");
                }

                // Save the modified workbook back to a memory stream
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    byte[] resultBytes = outputStream.ToArray();

                    // Write the result to a file
                    const string outputPath = "output.xlsx";
                    File.WriteAllBytes(outputPath, resultBytes);
                    Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
                }
            }
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
