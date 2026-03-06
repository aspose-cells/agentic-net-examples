using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class AdvancedXlsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data and a formula
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(0.5);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["C3"].PutValue(0.3);
        sheet.Cells["D1"].PutValue("Total");
        sheet.Cells["D2"].Formula = "B2*C2";
        sheet.Cells["D3"].Formula = "B3*C3";

        // Create XlsSaveOptions and configure desired properties
        XlsSaveOptions saveOptions = new XlsSaveOptions();
        saveOptions.MatchColor = true;               // Match palette colors for compatibility
        saveOptions.ClearData = false;               // Keep workbook data after saving
        saveOptions.ValidateMergedAreas = true;      // Validate merged cells before saving
        saveOptions.RefreshChartCache = true;        // Refresh chart cache (if any)
        saveOptions.UpdateSmartArt = false;          // Do not update SmartArt

        // Save the workbook as an Excel 97-2003 .xls file using the options
        string xlsPath = "AdvancedDemo.xls";
        workbook.Save(xlsPath, saveOptions);

        // Convert the saved .xls file to PDF using the ConversionUtility
        string pdfPath = "AdvancedDemo.pdf";
        ConversionUtility.Convert(xlsPath, pdfPath);

        // Save the same workbook as SpreadsheetML 2003 XML with XLS limits enforced
        SpreadsheetML2003SaveOptions xmlOptions = new SpreadsheetML2003SaveOptions();
        xmlOptions.LimitAsXls = true; // Enforce max rows (65535) and columns (255) of XLS format
        string xmlPath = "AdvancedDemo.xml";
        workbook.Save(xmlPath, xmlOptions);

        // Release resources
        workbook.Dispose();

        Console.WriteLine("Advanced XLS operations completed successfully.");
    }
}