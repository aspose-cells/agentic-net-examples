using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Load the existing ODS workbook
        Workbook workbook = new Workbook("input.ods");

        // Configure ODS save options to use ODF version 1.1 (strict)
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf11;

        // Save the workbook with ODF 1.1 compliance
        workbook.Save("output_strict.ods", saveOptions);
    }
}