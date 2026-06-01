using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class OdsToPdfConverter
{
    static void Main()
    {
        // Source ODS file path
        string sourcePath = "input.ods";

        // Destination PDF file path
        string destPath = "output.pdf";

        // Load the ODS workbook with default load options
        OdsLoadOptions loadOptions = new OdsLoadOptions();
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Save the workbook as PDF using default PDF save options
        workbook.Save(destPath, SaveFormat.Pdf);
    }
}