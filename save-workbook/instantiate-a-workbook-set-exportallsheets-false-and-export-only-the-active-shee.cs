using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ExportActiveSheetToXml
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a second worksheet for demonstration
        workbook.Worksheets.Add("Sheet2");

        // Fill data in both worksheets
        workbook.Worksheets[0].Cells["A1"].PutValue("Active Sheet Data");
        workbook.Worksheets[1].Cells["A1"].PutValue("Other Sheet Data");

        // Set the first worksheet as the active sheet
        workbook.Worksheets.ActiveSheetIndex = 0;

        // Set ExportAllSheets to false (property of TxtSaveOptions)
        TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
        txtOptions.ExportAllSheets = false;

        // Configure XmlSaveOptions to export only the active worksheet
        XmlSaveOptions xmlOptions = new XmlSaveOptions();
        xmlOptions.SheetIndexes = new int[] { workbook.Worksheets.ActiveSheetIndex };

        // Save the workbook as XML, exporting only the active sheet
        workbook.Save("ActiveSheetOnly.xml", xmlOptions);
    }
}