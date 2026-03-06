using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;

class AdvancedXlsxDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet and add sample data
        Worksheet ws = workbook.Worksheets[0];
        ws.Name = "Data";

        ws.Cells["A1"].PutValue("Product");
        ws.Cells["B1"].PutValue("Quantity");
        ws.Cells["C1"].PutValue("Price");
        ws.Cells["A2"].PutValue("Apple");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["C2"].PutValue(0.5);
        ws.Cells["A3"].PutValue("Banana");
        ws.Cells["B3"].PutValue(20);
        ws.Cells["C3"].PutValue(0.3);

        // Add a simple formula
        ws.Cells["D1"].PutValue("Total");
        ws.Cells["D2"].Formula = "B2*C2";
        ws.Cells["D3"].Formula = "B3*C3";

        // Save as XLSX using OoxmlSaveOptions (compression and cell name export)
        OoxmlSaveOptions ooxmlOptions = new OoxmlSaveOptions(); // rule
        ooxmlOptions.ExportCellName = true; // example property
        string xlsxPath = "AdvancedDemo.xlsx";
        workbook.Save(xlsxPath, ooxmlOptions); // rule

        // Save as XLS using XlsSaveOptions (match color to palette)
        XlsSaveOptions xlsOptions = new XlsSaveOptions(); // rule
        xlsOptions.MatchColor = true;
        string xlsPath = "AdvancedDemo.xls";
        workbook.Save(xlsPath, xlsOptions); // rule

        // Convert the generated XLSX file to PDF using ConversionUtility
        string pdfPath = "AdvancedDemo.pdf";
        ConversionUtility.Convert(xlsxPath, pdfPath); // rule

        // Detect the file format of the saved XLSX file
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(xlsxPath); // rule
        Console.WriteLine($"Detected format: {formatInfo.FileFormatType}, Encrypted: {formatInfo.IsEncrypted}");

        // Attempt to export an XML map (placeholder name "SampleMap")
        try
        {
            workbook.ExportXml("SampleMap", "ExportedSample.xml"); // rule
            Console.WriteLine("ExportXml completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ExportXml failed: {ex.Message}");
        }

        // Dispose the workbook
        workbook.Dispose();
    }
}