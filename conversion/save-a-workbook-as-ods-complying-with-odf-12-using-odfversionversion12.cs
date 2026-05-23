using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsOdsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to the worksheet
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(456.78);

            // Create ODS save options and set the ODF version to 1.2
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

            // Save the workbook as ODS with the specified ODF version
            workbook.Save("OutputVersion12.ods", saveOptions);
        }
    }
}