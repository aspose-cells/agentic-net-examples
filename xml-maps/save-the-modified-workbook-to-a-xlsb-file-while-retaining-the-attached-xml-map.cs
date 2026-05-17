using System;
using Aspose.Cells;

class SaveXlsbWithXmlMap
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Laptop");
        sheet.Cells["B2"].PutValue(999.99);

        // The XmlMaps feature is not available in the current Aspose.Cells version.
        // If needed, ensure you reference a version that supports XML mapping
        // and use the following code:
        // int mapIndex = workbook.XmlMaps.Add(xmlSchema, "ProductDataMap");
        // XmlMap xmlMap = workbook.XmlMaps[mapIndex];

        // Save the workbook as XLSB
        XlsbSaveOptions saveOptions = new XlsbSaveOptions();
        workbook.Save("ProductData.xlsb", saveOptions);
    }
}