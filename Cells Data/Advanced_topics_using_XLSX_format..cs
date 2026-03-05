using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;

namespace AdvancedXlsxDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Create a new workbook and add sample data
            Workbook workbook = new Workbook(); // create workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate header
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Populate some rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // 2. Set OOXML compliance to ISO/IEC 29500:2008 Strict
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // 3. Save the workbook as XLSX using OoxmlSaveOptions with compression
            OoxmlSaveOptions ooxmlOptions = new OoxmlSaveOptions(); // create OoxmlSaveOptions
            ooxmlOptions.CompressionType = OoxmlCompressionType.Level6; // enable high compression
            string xlsxPath = "AdvancedDemo.xlsx";
            workbook.Save(xlsxPath, ooxmlOptions); // save workbook

            Console.WriteLine($"Workbook saved as XLSX to '{xlsxPath}' with strict OOXML compliance.");

            // 4. Convert the saved XLSX file to PDF using ConversionUtility
            string pdfPath = "AdvancedDemo.pdf";
            ConversionUtility.Convert(xlsxPath, pdfPath); // convert to PDF
            Console.WriteLine($"Workbook converted to PDF at '{pdfPath}'.");

            // 5. Detect the file format of the generated XLSX file
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(xlsxPath);
            Console.WriteLine($"Detected format: {formatInfo.FileFormatType}, Encrypted: {formatInfo.IsEncrypted}");

            // 6. Save the same workbook as XLS (Excel 97-2003) using XlsSaveOptions
            XlsSaveOptions xlsOptions = new XlsSaveOptions(); // create XlsSaveOptions
            xlsOptions.MatchColor = true; // example option
            string xlsPath = "AdvancedDemo.xls";
            workbook.Save(xlsPath, xlsOptions);
            Console.WriteLine($"Workbook also saved as XLS to '{xlsPath}'.");

            // 7. Save the workbook as XLSB using XlsbSaveOptions
            XlsbSaveOptions xlsbOptions = new XlsbSaveOptions(); // create XlsbSaveOptions
            xlsbOptions.ExportAllColumnIndexes = true; // example option
            string xlsbPath = "AdvancedDemo.xlsb";
            workbook.Save(xlsbPath, xlsbOptions);
            Console.WriteLine($"Workbook also saved as XLSB to '{xlsbPath}'.");

            // 8. Export XML data linked by an XML map (placeholder example)
            // Note: An XML map must be defined in the workbook before exporting.
            // For demonstration, we assume a map named "SampleMap" exists.
            try
            {
                string xmlMapName = "SampleMap";
                string xmlExportPath = "ExportedData.xml";
                workbook.ExportXml(xmlMapName, xmlExportPath);
                Console.WriteLine($"XML data exported to '{xmlExportPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ExportXml failed (expected if no XML map defined): {ex.Message}");
            }

            // Cleanup
            workbook.Dispose();
        }
    }
}