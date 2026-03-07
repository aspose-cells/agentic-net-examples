using System;
using System.IO;
using Aspose.Cells;

namespace SpreadsheetMLToJsonDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string inputPath = Path.Combine(baseDir, "sample_spreadsheetml.xml");
            string outputPath = Path.Combine(baseDir, "output.json");

            EnsureSampleSpreadsheetMLExists(inputPath);
            ConvertSpreadsheetMLToJson(inputPath, outputPath);

            Console.WriteLine($"Conversion completed. JSON saved to: {outputPath}");
        }

        static void EnsureSampleSpreadsheetMLExists(string path)
        {
            if (File.Exists(path))
                return;

            string sampleXml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet""
          xmlns:o=""urn:schemas-microsoft-com:office:office""
          xmlns:x=""urn:schemas-microsoft-com:office:excel""
          xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet""
          xmlns:html=""http://www.w3.org/TR/REC-html40"">
  <Worksheet ss:Name=""Sheet1"">
    <Table>
      <Row>
        <Cell><Data ss:Type=""String"">Hello</Data></Cell>
        <Cell><Data ss:Type=""Number"">123</Data></Cell>
      </Row>
    </Table>
  </Worksheet>
</Workbook>";
            File.WriteAllText(path, sampleXml);
        }

        static void ConvertSpreadsheetMLToJson(string inputFile, string outputFile)
        {
            LoadOptions loadOptions = new LoadOptions(LoadFormat.SpreadsheetML);
            Workbook workbook = new Workbook(inputFile, loadOptions);

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                AlwaysExportAsJsonObject = true,
                Indent = "  "
            };

            workbook.Save(outputFile, jsonOptions);
        }
    }
}