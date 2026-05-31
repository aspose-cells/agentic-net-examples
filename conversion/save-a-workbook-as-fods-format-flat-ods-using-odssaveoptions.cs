using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456.78);

        // Create OdsSaveOptions specifying the FODS format
        OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.Fods);
        // Optional: set generator and ODF version
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

        // Save the workbook as a Flat ODS file
        workbook.Save("output.fods", saveOptions);
    }
}