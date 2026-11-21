// Title: Serialize custom XML parts to string and log them from an Excel workbook using Aspose.Cells for .NET
// Common Searches: c# serialize custom XML part to string with Aspose.Cells | how to log custom XML part content from an .xlsx using Aspose.Cells | read custom XML parts from workbook when GetXml method is unavailable Aspose.Cells | retrieve custom XML part Id or Name in Aspose.Cells C# | debug Excel custom XML parts using Aspose.Cells reflection
// Tags: Aspose.Cells GetXml custom XML part | Aspose.Cells XmlData fallback reading | Aspose.Cells read custom XML from .xlsx | Aspose.Cells output custom XML to console | Aspose.Cells custom part Id extraction

using Aspose.Cells;
using System;
using System.IO;
using System.Text;
using System.Reflection;

// The example loads an Excel workbook with Aspose.Cells, iterates through each CustomXmlPart, uses reflection to invoke the GetXml method or falls back to the XmlData property, converts the XML bytes to a UTF‑8 string, obtains the part's Id or Name, and writes both the identifier and the XML content to the console while handling errors at both the part and workbook levels.
class Program
{
    static void Main()
    {
        string filePath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Iterate through all custom XML parts in the workbook
            foreach (var customXml in workbook.CustomXmlParts)
            {
                try
                {
                    // Try to obtain the XML content using the available API
                    string xmlContent = string.Empty;
                    MethodInfo getXmlMethod = customXml.GetType().GetMethod("GetXml", BindingFlags.Public | BindingFlags.Instance);
                    if (getXmlMethod != null)
                    {
                        xmlContent = (string)getXmlMethod.Invoke(customXml, null);
                    }
                    else
                    {
                        // Fallback: read raw XML data if GetXml is not present
                        PropertyInfo xmlDataProp = customXml.GetType().GetProperty("XmlData", BindingFlags.Public | BindingFlags.Instance);
                        if (xmlDataProp != null)
                        {
                            byte[] data = (byte[])xmlDataProp.GetValue(customXml);
                            xmlContent = Encoding.UTF8.GetString(data);
                        }
                    }

                    // Retrieve an identifier for the part (Id or Name)
                    string partId = string.Empty;
                    PropertyInfo idProp = customXml.GetType().GetProperty("Id", BindingFlags.Public | BindingFlags.Instance);
                    if (idProp != null)
                    {
                        partId = idProp.GetValue(customXml)?.ToString();
                    }
                    else
                    {
                        PropertyInfo nameProp = customXml.GetType().GetProperty("Name", BindingFlags.Public | BindingFlags.Instance);
                        if (nameProp != null)
                        {
                            partId = nameProp.GetValue(customXml)?.ToString();
                        }
                    }

                    Console.WriteLine($"Custom XML Part ID/Name: {partId}");
                    Console.WriteLine(xmlContent);
                }
                catch (Exception exPart)
                {
                    // Handle errors for individual custom XML parts
                    Console.WriteLine($"Error processing a custom XML part: {exPart.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
