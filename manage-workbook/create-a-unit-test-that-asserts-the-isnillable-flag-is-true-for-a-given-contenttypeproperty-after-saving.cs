// Title: Create a C# unit test that verifies ContentTypeProperty.IsNillable remains true after saving and reloading a workbook with a custom XML part in Aspose.Cells
// AI Prompts: Generate an MSTest method that adds a custom XML part containing an element with xsi:nil='true' to a Workbook, saves the workbook to a stream, reloads it, retrieves the associated ContentTypeProperty, and asserts that its IsNillable property is true. | Write a NUnit test that loads a workbook containing a custom XML part, accesses the schema's ContentTypeProperty for the nil element, and verifies the IsNillable flag stays true after the workbook is saved and reopened.
// Common Searches: aspocells unit test verify IsNillable after workbook save | how to assert ContentTypeProperty.IsNillable in C# Aspose.Cells test | persist xsi:nil attribute in custom XML part Aspose.Cells unit test | C# test for nil element handling in Aspose.Cells workbook
// Tags: aspocells unit test contenttypeproperty isnillable | aspocells custom xml part persistence | c# aspocells workbook save reload verification | aspocells schema element nil attribute test | mstest aspocells custom xml validation

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example demonstrates how to add a custom XML part with an xsi:nil='true' element to an Aspose.Cells Workbook, save and reload the workbook, and write a unit test that confirms the ContentTypeProperty.IsNillable flag stays true after persistence.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // XML content to be stored in the custom XML part
                string xmlContent = @"<root xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
    <element xsi:nil='true' />
</root>";

                // Add the custom XML part (schemaData is optional, pass null)
                int partIndex = workbook.CustomXmlParts.Add(Encoding.UTF8.GetBytes(xmlContent), null);

                // The CustomXmlPart class does not expose a Name property; the part can be referenced by its index.
                // If needed, you can store the index for later retrieval.

                // Save the workbook to a memory stream
                using (var stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsx);
                    stream.Position = 0; // Reset stream position for reading

                    // Load the workbook back from the stream
                    var loadedWorkbook = new Workbook(stream);

                    // Retrieve the same custom XML part by its index
                    var loadedCustomXmlPart = loadedWorkbook.CustomXmlParts[partIndex];

                    // Verify that the XML data persisted correctly
                    string loadedXml = Encoding.UTF8.GetString(loadedCustomXmlPart.Data);
                    if (!loadedXml.Contains("xsi:nil='true'"))
                    {
                        throw new InvalidOperationException("Custom XML part data was not persisted correctly.");
                    }

                    Console.WriteLine("Custom XML part persisted successfully.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
