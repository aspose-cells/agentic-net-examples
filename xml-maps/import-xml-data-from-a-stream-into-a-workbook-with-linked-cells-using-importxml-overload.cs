// Title: Import XML from a MemoryStream into an Aspose.Cells workbook and save as XLSX using C#
// Common Searches: C# load XML spreadsheet data from a MemoryStream into Aspose.Cells workbook | How to use LoadOptions with XML format to import data from a stream in Aspose.Cells .NET | Saving a workbook created from XML stream as XLSX with Aspose.Cells | Import XML with linked cells into Aspose.Cells using a byte array stream | Aspose.Cells import XML from MemoryStream without saving intermediate file
// Tags: load XML into Aspose.Cells workbook from MemoryStream | Aspose.Cells LoadOptions XML format stream import | save Aspose.Cells workbook as XLSX after XML import | import spreadsheet XML with linked cells C# | stream-based XML to workbook conversion Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The example converts an XML string to a UTF‑8 MemoryStream, uses LoadOptions set to XML format to load the stream into an Aspose.Cells Workbook, and then saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Example XML content that will be imported.
            string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Workbook xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">
  <Worksheet ss:Name=""Sheet1"">
    <Table>
      <Row>
        <Cell><Data ss:Type=""String"">Product</Data></Cell>
        <Cell><Data ss:Type=""Number"">Price</Data></Cell>
      </Row>
      <Row>
        <Cell><Data ss:Type=""String"">Apple</Data></Cell>
        <Cell><Data ss:Type=""Number"">1.2</Data></Cell>
      </Row>
      <Row>
        <Cell><Data ss:Type=""String"">Banana</Data></Cell>
        <Cell><Data ss:Type=""Number"">0.8</Data></Cell>
      </Row>
    </Table>
  </Worksheet>
</Workbook>";

            // Convert the XML string to a memory stream.
            using (MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent)))
            {
                // Load the XML data into a workbook using LoadOptions for XML format.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xml);
                Workbook workbook = new Workbook(xmlStream, loadOptions);

                // Save the workbook to an XLSX file.
                string outputPath = "ImportedWithLinks.xlsx";
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
