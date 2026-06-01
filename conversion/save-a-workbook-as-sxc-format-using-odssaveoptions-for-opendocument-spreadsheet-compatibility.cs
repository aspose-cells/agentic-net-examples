using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

class SaveAsSxc
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and add some data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Create OdsSaveOptions specifying the SXC format
        OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.Sxc);
        // Set the generator type for better OpenDocument compatibility (optional)
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Save the workbook as a StarOffice Calc Spreadsheet (.sxc) using the options
        workbook.Save("output.sxc", saveOptions);
    }
}