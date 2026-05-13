using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class Program
{
    static void Main()
    {
        // Path to the source FODS file
        string sourceFile = "input.fods";

        // Path for the resulting ODS file
        string destFile = "output.ods";

        // Load the FODS workbook using OdsLoadOptions (default constructor)
        OdsLoadOptions loadOptions = new OdsLoadOptions();
        Workbook workbook = new Workbook(sourceFile, loadOptions);

        // Change the workbook's default font
        workbook.DefaultStyle.Font.Name = "Calibri";
        workbook.DefaultStyle.Font.Size = 11;

        // Prepare ODS save options (optional: set generator type)
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Save the workbook as an ODS file using the save options
        workbook.Save(destFile, saveOptions);
    }
}