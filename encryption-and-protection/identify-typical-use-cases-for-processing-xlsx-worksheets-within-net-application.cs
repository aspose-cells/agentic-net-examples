using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // 1. Create a new workbook, add data, apply simple formatting, and save it.
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Data";
        ws.Cells["A1"].PutValue("Product");
        ws.Cells["B1"].PutValue("Price");
        ws.Cells["A2"].PutValue("Laptop");
        ws.Cells["B2"].PutValue(999.99);
        ws.Cells["A3"].PutValue("Phone");
        ws.Cells["B3"].PutValue(699.99);

        // Apply bold style to header row.
        Style headerStyle = wb.CreateStyle();
        headerStyle.Font.IsBold = true;
        StyleFlag flag = new StyleFlag { All = true };
        ws.Cells.CreateRange("A1:B1").ApplyStyle(headerStyle, flag);

        // Auto‑fit columns to content.
        ws.AutoFitColumns();

        // Save the workbook to disk.
        wb.Save("CreatedWorkbook.xlsx");

        // 2. Load the workbook, read a cell value, modify data, add a new sheet, and save.
        Workbook loadedWb = new Workbook("CreatedWorkbook.xlsx");
        Worksheet loadedWs = loadedWb.Worksheets["Data"];

        // Read values from cells.
        string product = loadedWs.Cells["A2"].StringValue;
        double price = loadedWs.Cells["B2"].DoubleValue;
        Console.WriteLine($"Read from file: {product} costs {price}");

        // Apply a 10% discount to the price.
        loadedWs.Cells["B2"].PutValue(price * 0.9);

        // Add a summary worksheet.
        int summaryIndex = loadedWb.Worksheets.Add();
        Worksheet summaryWs = loadedWb.Worksheets[summaryIndex];
        summaryWs.Name = "Summary";
        summaryWs.Cells["A1"].PutValue("Discounted Price");
        summaryWs.Cells["B1"].PutValue(loadedWs.Cells["B2"].DoubleValue);

        // Save the modified workbook.
        loadedWb.Save("ModifiedWorkbook.xlsx");

        // 3. Copy the modified workbook into a new workbook instance.
        Workbook copyWb = new Workbook();
        copyWb.Copy(loadedWb);
        copyWb.Save("CopyWorkbook.xlsx");

        // 4. Convert the modified workbook to PDF using the ConversionUtility.
        ConversionUtility.Convert("ModifiedWorkbook.xlsx", "Report.pdf");

        // 5. Import XML data from a stream into a new workbook.
        string xmlContent = @"<Products>
                                <Product>
                                    <Name>Tablet</Name>
                                    <Price>399.99</Price>
                                </Product>
                              </Products>";
        using (MemoryStream xmlStream = new MemoryStream())
        {
            using (StreamWriter writer = new StreamWriter(xmlStream, Encoding.UTF8, 1024, true))
            {
                writer.Write(xmlContent);
                writer.Flush();
                xmlStream.Position = 0;
            }

            Workbook xmlWb = new Workbook();
            xmlWb.ImportXml(xmlStream, "Sheet1", 0, 0);
            xmlWb.Save("XmlImported.xlsx");
        }

        // 6. Export worksheet data to XML using an XML map (placeholder example).
        // wb.ExportXml("SampleMap", "ExportedData.xml");
    }
}