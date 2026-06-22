using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("ODF 1.3 Demo");

        // Create ODS save options
        OdsSaveOptions saveOptions = new OdsSaveOptions();

        // Set the ODF version to 1.3 (ODF 1.3 compliance)
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf13;

        // Save the workbook as an ODS file with the specified ODF version
        workbook.Save("DemoOdf13.ods", saveOptions);
    }
}