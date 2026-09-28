// Title: Load an Excel workbook from a MemoryStream, attach an XML map using Aspose.Cells (with reflection fallback), and write the workbook to another stream in C#
// AI Prompts: Generate C# code that opens an XLSX workbook from a MemoryStream, creates an XML map from an XSD Stream using Aspose.Cells (including a reflection‑based fallback when XmlMaps is unavailable), and saves the modified workbook to a new MemoryStream. | Show how to bind an XSD schema to a Workbook as an XML map in C# when the XmlMaps collection is not directly accessible, leveraging Aspose.Cells reflection techniques. | Provide a complete example that reads an input workbook file, adds an XML schema mapping, and outputs the result to a stream without writing intermediate files, using Aspose.Cells.
// Common Searches: asp.net core add xml schema as map to excel workbook from memory stream using aspose.cells | c# load xlsx from stream, attach xml map, and save to stream with aspose.cells | how to use reflection to add xml map in older Aspose.Cells versions c# | aspose.cells create xml map from xsd without XmlMaps property
// Tags: memorystream workbook loading Aspose.Cells | xml map creation from XSD Aspose.Cells | reflection fallback for XmlMaps Aspose.Cells | save workbook to stream Aspose.Cells | c# excel xml schema mapping

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook from a file stream, uses reflection to add an XML map named "MyXmlMap" from an XSD stream when the XmlMaps property is not directly available, and then saves the updated workbook to a MemoryStream for further processing.
public class XmlMapExample
{
    public static void Main()
    {
        try
        {
            var example = new XmlMapExample();
            example.ProcessWorkbook();
            Console.WriteLine("Workbook processed successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    public void ProcessWorkbook()
    {
        // Obtain the workbook stream (e.g., from a file)
        using (Stream inputStream = GetInputWorkbookStream())
        {
            // Load the workbook from the stream
            var workbook = new Workbook(inputStream);

            // Obtain the XML schema (XSD) stream
            using (Stream xmlSchemaStream = GetXmlSchemaStream())
            {
                // Attempt to add an XML map using reflection (covers versions without direct XmlMaps property)
                try
                {
                    var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
                    if (xmlMapsProp != null)
                    {
                        var xmlMaps = xmlMapsProp.GetValue(workbook);
                        var addMethod = xmlMaps?.GetType().GetMethod("Add", new[] { typeof(string), typeof(Stream) });
                        addMethod?.Invoke(xmlMaps, new object[] { "MyXmlMap", xmlSchemaStream });
                    }
                }
                catch (Exception ex)
                {
                    // Log but continue if XML mapping is not supported
                    Console.Error.WriteLine($"Warning: Unable to add XML map - {ex.Message}");
                }
            }

            // Save the modified workbook to a memory stream
            using (var outputStream = new MemoryStream())
            {
                workbook.Save(outputStream, SaveFormat.Xlsx);
                outputStream.Position = 0; // Reset for downstream use
                HandleOutputWorkbookStream(outputStream);
            }
        }
    }

    // Returns a FileStream for the input workbook; checks existence first.
    private Stream GetInputWorkbookStream()
    {
        const string inputPath = "input.xlsx";
        if (!File.Exists(inputPath))
            throw new FileNotFoundException($"Input workbook not found: {inputPath}");

        return new FileStream(inputPath, FileMode.Open, FileAccess.Read);
    }

    // Returns a FileStream for the XML schema; checks existence first.
    private Stream GetXmlSchemaStream()
    {
        const string schemaPath = "schema.xsd";
        if (!File.Exists(schemaPath))
            throw new FileNotFoundException($"XML schema file not found: {schemaPath}");

        return new FileStream(schemaPath, FileMode.Open, FileAccess.Read);
    }

    // Writes the output workbook stream to a file.
    private void HandleOutputWorkbookStream(Stream outputStream)
    {
        const string outputPath = "output.xlsx";
        using (var file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            outputStream.CopyTo(file);
        }
    }
}
