using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class AdvancedXlsxHandler
{
    static void Main()
    {
        // Paths for demonstration files
        string inputPath = "input.xlsx";
        string xmlPath = "data.xml";
        string outputXlsx = "output_optimized.xlsx";
        string outputPdf = "output.pdf";
        string outputXlsb = "output.xlsb";
        string outputXml2003 = "output_2003.xml";

        // Ensure sample workbook and XML exist
        EnsureSampleWorkbook(inputPath);
        EnsureSampleXml(xmlPath);

        // Detect the format of the input file
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(inputPath);
        Console.WriteLine($"Detected format: {formatInfo.FileFormatType}");

        // Load workbook with memory‑optimized LoadOptions
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.MemorySetting = MemorySetting.MemoryPreference;
        Workbook wb = new Workbook(inputPath, loadOptions);
        Console.WriteLine($"Loaded workbook with {wb.Worksheets.Count} worksheet(s).");

        // Import XML data into the first worksheet starting at cell A1
        wb.ImportXml(xmlPath, wb.Worksheets[0].Name, 0, 0);
        Console.WriteLine("XML data imported.");

        // Save with OoxmlSaveOptions for performance (compression, no cell names)
        OoxmlSaveOptions ooxmlOptions = new OoxmlSaveOptions();
        ooxmlOptions.CompressionType = OoxmlCompressionType.Level6;
        ooxmlOptions.ExportCellName = false;          // improves speed and reduces size
        ooxmlOptions.EnableZip64 = false;
        ooxmlOptions.UpdateZoom = true;
        wb.Save(outputXlsx, ooxmlOptions);
        Console.WriteLine($"Workbook saved to {outputXlsx} with OoxmlSaveOptions.");

        // Save to a MemoryStream as an XLS file and reload from the stream
        MemoryStream xlsStream = wb.SaveToStream();
        Workbook wbFromStream = new Workbook(xlsStream);
        Console.WriteLine("Workbook reloaded from MemoryStream.");

        // Convert the optimized XLSX to PDF using ConversionUtility
        ConversionUtility.Convert(outputXlsx, outputPdf);
        Console.WriteLine($"Workbook converted to PDF: {outputPdf}");

        // Save as Xlsb using XlsbSaveOptions
        XlsbSaveOptions xlsbOptions = new XlsbSaveOptions();
        xlsbOptions.ExportAllColumnIndexes = true;
        wb.Save(outputXlsb, xlsbOptions);
        Console.WriteLine($"Workbook saved as Xlsb: {outputXlsb}");

        // Save as SpreadsheetML2003 with XLS limits (for legacy compatibility)
        SpreadsheetML2003SaveOptions xml2003Options = new SpreadsheetML2003SaveOptions();
        xml2003Options.LimitAsXls = true;               // enforce 65535 rows / 255 columns limit
        xml2003Options.IsIndentedFormatting = true;
        wb.Save(outputXml2003, xml2003Options);
        Console.WriteLine($"Workbook saved as SpreadsheetML2003: {outputXml2003}");

        // Save as XLS with encryption and clear data after saving
        XlsSaveOptions encOptions = new XlsSaveOptions();
        encOptions.EncryptDocumentProperties = true;    // encrypt document properties
        encOptions.ClearData = true;                    // workbook cleared after save
        wb.Save("encrypted.xls", encOptions);
        Console.WriteLine("Workbook saved with encryption and cleared data.");

        // Clean up resources
        xlsStream.Dispose();
    }

    // Creates a simple workbook if the specified file does not exist
    static void EnsureSampleWorkbook(string path)
    {
        if (!File.Exists(path))
        {
            Workbook temp = new Workbook();
            Worksheet ws = temp.Worksheets[0];
            ws.Cells["A1"].PutValue("Sample");
            ws.Cells["B1"].PutValue(DateTime.Now);
            temp.Save(path, SaveFormat.Xlsx);
        }
    }

    // Creates a simple XML file if the specified file does not exist
    static void EnsureSampleXml(string path)
    {
        if (!File.Exists(path))
        {
            string xml = @"<Products>
    <Product><Name>Item1</Name><Price>10.5</Price></Product>
    <Product><Name>Item2</Name><Price>20.0</Price></Product>
</Products>";
            File.WriteAllText(path, xml);
        }
    }
}