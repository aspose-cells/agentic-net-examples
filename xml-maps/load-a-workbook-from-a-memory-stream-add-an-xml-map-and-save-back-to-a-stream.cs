using System;
using System.IO;
using Aspose.Cells;

public class XmlMapExample
{
    public static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Prepare an input Excel file in memory (for demo purposes)
            // ------------------------------------------------------------
            using (MemoryStream inputStream = new MemoryStream())
            {
                // Create a simple workbook and save it to a MemoryStream.
                using (Workbook tempWorkbook = new Workbook())
                {
                    Worksheet sheet = tempWorkbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Sample");
                    sheet.Cells["B1"].PutValue(123);
                    tempWorkbook.Save(inputStream, SaveFormat.Xlsx);
                }

                // Reset the position so it can be read from the beginning
                inputStream.Position = 0;

                // ------------------------------------------------------------
                // 2. Load the workbook from the memory stream
                // ------------------------------------------------------------
                using (Workbook workbook = new Workbook(inputStream))
                {
                    // ------------------------------------------------------------
                    // 3. Add an XML map to the workbook
                    // ------------------------------------------------------------
                    // Define a simple XML schema (XSD) as a string.
                    string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                            <xs:element name='Root'>
                                                <xs:complexType>
                                                    <xs:sequence>
                                                        <xs:element name='Item' type='xs:string'/>
                                                    </xs:sequence>
                                                </xs:complexType>
                                            </xs:element>
                                        </xs:schema>";

                    // Write the schema to a temporary file because XmlMaps.Add expects a file path.
                    string tempSchemaPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
                    File.WriteAllText(tempSchemaPath, xmlSchema);

                    // Ensure the file exists before adding.
                    if (File.Exists(tempSchemaPath))
                    {
                        int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);
                        XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                        xmlMap.Name = "DemoMap";
                    }
                    else
                    {
                        throw new FileNotFoundException("Temporary XML schema file was not created.", tempSchemaPath);
                    }

                    // Clean up the temporary schema file.
                    File.Delete(tempSchemaPath);

                    // ------------------------------------------------------------
                    // 4. Save the modified workbook back to a memory stream
                    // ------------------------------------------------------------
                    using (MemoryStream outputStream = workbook.SaveToStream())
                    {
                        // Write the stream to a file on disk.
                        using (FileStream file = new FileStream("WorkbookWithXmlMap.xls", FileMode.Create, FileAccess.Write))
                        {
                            outputStream.WriteTo(file);
                        }
                    }
                }
            }

            Console.WriteLine("Workbook loaded, XML map added, and saved to 'WorkbookWithXmlMap.xls'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}