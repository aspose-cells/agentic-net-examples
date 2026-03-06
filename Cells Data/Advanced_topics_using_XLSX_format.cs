using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;

namespace AsposeCellsAdvancedXlsxDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "AdvancedDemo";

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(0.3);

            // 2. Set OOXML compliance level (advanced workbook setting)
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // 3. Save the workbook as XLSX with OOXML save options (save rule)
            OoxmlSaveOptions ooxmlOptions = new OoxmlSaveOptions();
            // Example: enable compression (property exists in Aspose.Cells)
            ooxmlOptions.CompressionType = OoxmlCompressionType.Level6;
            string xlsxPath = "AdvancedDemo.xlsx";
            workbook.Save(xlsxPath, ooxmlOptions);

            Console.WriteLine($"Workbook saved to '{xlsxPath}' with strict OOXML compliance.");

            // 4. Convert the saved XLSX to PDF using ConversionUtility (conversion rule)
            string pdfPath = "AdvancedDemo.pdf";
            ConversionUtility.Convert(xlsxPath, pdfPath);
            Console.WriteLine($"Workbook converted to PDF at '{pdfPath}'.");

            // 5. Detect the file format of the XLSX using a MemoryStream (detect rule)
            using (MemoryStream ms = new MemoryStream())
            {
                // Load the workbook again (load rule) and save to stream in XLSX format
                Workbook wbForDetect = new Workbook(xlsxPath);
                wbForDetect.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0; // Reset stream position for detection

                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(ms);
                Console.WriteLine($"Detected format: {formatInfo.FileFormatType}, Encrypted: {formatInfo.IsEncrypted}");
            }

            // 6. Export XML data linked by an XML map (export rule)
            // Note: For demonstration, we assume an XML map named "SampleMap" exists.
            // In a real scenario, the map should be created or loaded from the workbook.
            try
            {
                string xmlMapName = "SampleMap"; // placeholder name
                string xmlExportPath = "ExportedData.xml";
                workbook.ExportXml(xmlMapName, xmlExportPath);
                Console.WriteLine($"XML data exported to '{xmlExportPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ExportXml failed: {ex.Message}");
            }

            // Cleanup
            workbook.Dispose();
        }
    }
}