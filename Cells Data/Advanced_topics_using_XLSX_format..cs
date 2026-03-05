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
            // ------------------------------------------------------------
            // 1. Create a new workbook and add some sample data
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];              // get the first worksheet

            // Populate header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            // Populate a few data rows
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(85);
            sheet.Cells["C3"].PutValue(0.3);

            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(60);
            sheet.Cells["C4"].PutValue(1.2);

            // ------------------------------------------------------------
            // 2. Set OOXML compliance to ISO/IEC 29500:2008 Strict
            // ------------------------------------------------------------
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // ------------------------------------------------------------
            // 3. Save the workbook as XLSX using OoxmlSaveOptions
            //    (demonstrates setting compression level)
            // ------------------------------------------------------------
            OoxmlSaveOptions ooxmlOptions = new OoxmlSaveOptions();          // create OOXML save options
            ooxmlOptions.CompressionType = OoxmlCompressionType.Level6;     // set high compression
            string xlsxPath = "AdvancedDemo.xlsx";
            workbook.Save(xlsxPath, ooxmlOptions);                           // save with options

            Console.WriteLine($"Workbook saved as XLSX: {xlsxPath}");

            // ------------------------------------------------------------
            // 4. Convert the saved XLSX file to PDF using ConversionUtility
            // ------------------------------------------------------------
            string pdfPath = "AdvancedDemo.pdf";
            ConversionUtility.Convert(xlsxPath, pdfPath);                    // convert to PDF

            Console.WriteLine($"Workbook converted to PDF: {pdfPath}");

            // ------------------------------------------------------------
            // 5. Detect the file format of the generated PDF
            // ------------------------------------------------------------
            FileFormatInfo pdfInfo = FileFormatUtil.DetectFileFormat(pdfPath);
            Console.WriteLine($"Detected format for '{pdfPath}': {pdfInfo.FileFormatType}");
            Console.WriteLine($"Is encrypted: {pdfInfo.IsEncrypted}");

            // ------------------------------------------------------------
            // 6. Save the same workbook as legacy XLS using XlsSaveOptions
            //    (demonstrates matching color palette)
            // ------------------------------------------------------------
            XlsSaveOptions xlsOptions = new XlsSaveOptions();               // create XLS save options
            xlsOptions.MatchColor = true;                                   // enable color matching
            string xlsPath = "AdvancedDemo.xls";
            workbook.Save(xlsPath, xlsOptions);                             // save as XLS

            Console.WriteLine($"Workbook saved as legacy XLS: {xlsPath}");

            // ------------------------------------------------------------
            // 7. Export XML data linked by an XML map (placeholder example)
            //    Note: In a real scenario the workbook must contain an XML map.
            // ------------------------------------------------------------
            try
            {
                string xmlMapName = "SampleMap"; // replace with actual map name if present
                string xmlExportPath = "ExportedMap.xml";
                workbook.ExportXml(xmlMapName, xmlExportPath);
                Console.WriteLine($"XML map '{xmlMapName}' exported to: {xmlExportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ExportXml failed (likely no XML map present): {ex.Message}");
            }

            // ------------------------------------------------------------
            // Cleanup
            // ------------------------------------------------------------
            workbook.Dispose();
        }
    }
}