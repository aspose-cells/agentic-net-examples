using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;

class XlsxHandlingDemo
{
    static void Main()
    {
        // 1. Create a new workbook (default format is Xlsx)
        Workbook wb = new Workbook();

        // Access the default worksheet and rename it
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Data";

        // Write header and sample data
        ws.Cells["A1"].PutValue("Product");
        ws.Cells["B1"].PutValue("Price");
        ws.Cells["A2"].PutValue("Apple");
        ws.Cells["B2"].PutValue(1.5);
        ws.Cells["A3"].PutValue("Banana");
        ws.Cells["B3"].PutValue(0.75);

        // Apply a simple style to the header row
        Style headerStyle = wb.CreateStyle();
        headerStyle.Font.IsBold = true;
        StyleFlag flag = new StyleFlag { All = true };
        ws.Cells.CreateRange("A1:B1").ApplyStyle(headerStyle, flag);

        // Save the workbook to an Xlsx file
        string xlsxPath = "Demo.xlsx";
        wb.Save(xlsxPath, SaveFormat.Xlsx);

        // 2. Load the workbook from the file we just saved
        Workbook loadedWb = new Workbook(xlsxPath);
        Worksheet loadedWs = loadedWb.Worksheets["Data"];

        // Ensure any formulas are calculated
        loadedWb.CalculateFormula();

        // Add a new worksheet for summary information
        int summaryIndex = loadedWb.Worksheets.Add();
        Worksheet summaryWs = loadedWb.Worksheets[summaryIndex];
        summaryWs.Name = "Summary";

        // Write a label and a formula that sums the prices
        summaryWs.Cells["A1"].PutValue("Total Price");
        summaryWs.Cells["B1"].Formula = "=SUM(Data!B2:B3)";
        loadedWb.CalculateFormula();

        // Save the workbook to PDF using PdfSaveOptions
        string pdfPath = "Demo.pdf";
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        loadedWb.Save(pdfPath, pdfOptions);

        // 3. Convert a CSV file to Xlsx using the static ConversionUtility
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath, "Item,Qty\nPen,10\nPencil,20");
        string csvToXlsx = "Converted.xlsx";
        ConversionUtility.Convert(csvPath, csvToXlsx);

        // 4. Detect the file format of the CSV using FileFormatUtil
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(csvPath);
        Console.WriteLine($"Detected format for '{csvPath}': {formatInfo.FileFormatType}");

        // 5. Save a workbook to a memory stream and reload it
        MemoryStream stream = wb.SaveToStream();
        stream.Position = 0;
        Workbook fromStream = new Workbook(stream);
        // Add a new column in the in‑memory workbook
        fromStream.Worksheets[0].Cells["C1"].PutValue("Notes");
        fromStream.Worksheets[0].Cells["C2"].PutValue("Sample");

        // Save the final workbook
        string finalPath = "FinalDemo.xlsx";
        fromStream.Save(finalPath, SaveFormat.Xlsx);

        // Clean up temporary files used for conversion demo
        File.Delete(csvPath);
        File.Delete(csvToXlsx);

        Console.WriteLine("XLSX handling demo completed successfully.");
    }
}