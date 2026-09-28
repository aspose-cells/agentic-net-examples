// Title: Import XML spreadsheet from a MemoryStream into an Aspose.Cells Workbook while preserving linked cells (C#)
// AI Prompts: Generate C# code that creates a MemoryStream from an XML string and uses Aspose.Cells Workbook.ImportXml overload to load the data with linked cells intact. | Show how to load an XML spreadsheet via a stream into a Workbook, then save it as an XLSX file using Aspose.Cells in C#. | Provide a step‑by‑step example of preserving cell links when importing XML data from a stream into an Aspose.Cells workbook.
// Common Searches: aspose.cells importxml from memorystream c# preserving linked cells | c# load xml spreadsheet into workbook using stream with aspose.cells | how to convert xml spreadsheet to xlsx while keeping cell links in c# | example of Workbook.ImportXml overload with MemoryStream in Aspose.Cells | import xml data into Aspose.Cells workbook from a stream c# tutorial
// Tags: importxml overload memorystream c# | aspose.cells load xml spreadsheet | preserve linked cells aspose.cells | convert xml to xlsx using aspose.cells | c# memorystream xml workbook import

using System;
using System.IO;
using Aspose.Cells;

// The example creates an XML spreadsheet string, writes it into a MemoryStream, uses Aspose.Cells Workbook.ImportXml (or the stream constructor) to load the XML while keeping linked cells, and then saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Sample XML content to be imported
            string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Workbook xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">
  <Worksheet ss:Name=""Sheet1"">
    <Table>
      <Row>
        <Cell><Data ss:Type=""String"">Name</Data></Cell>
        <Cell><Data ss:Type=""String"">Age</Data></Cell>
      </Row>
      <Row>
        <Cell><Data ss:Type=""String"">John</Data></Cell>
        <Cell><Data ss:Type=""Number"">30</Data></Cell>
      </Row>
    </Table>
  </Worksheet>
</Workbook>";

            // Convert the XML string into a memory stream
            using (MemoryStream xmlStream = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(xmlStream))
                {
                    writer.Write(xmlContent);
                    writer.Flush();
                    xmlStream.Position = 0; // Reset stream position for reading
                }

                // Load the workbook from the XML stream
                Workbook workbook = new Workbook(xmlStream);

                // Save the workbook to a file
                string outputPath = "ImportedFromXml.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
