using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Load the FODS file using OdsLoadOptions (default constructor)
        OdsLoadOptions loadOptions = new OdsLoadOptions();
        Workbook workbook = new Workbook("input.fods", loadOptions);

        // Change the workbook's default font
        workbook.DefaultStyle.Font.Name = "Calibri";
        workbook.DefaultStyle.Font.Size = 11;

        // Prepare ODS save options (optional: set generator type)
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Save the workbook as an ODS file
        workbook.Save("output.ods", saveOptions);
    }
}