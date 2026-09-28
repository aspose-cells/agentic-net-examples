// Title: Saving a workbook to Flat ODS (FODS) with Aspose.Cells OdsSaveOptions in C#
// AI Prompts: Write C# code that creates a Workbook, populates cells, configures OdsSaveOptions for the LibreOffice generator and ODF 1.2 strict mode, and saves the result as a .fods file. | Show how to use Aspose.Cells OdsSaveOptions to export a workbook to Flat ODS, including setting GeneratorType and OdfStrictVersion before calling Workbook.Save.
// Common Searches: C# Aspose.Cells export workbook to flat ODS file with custom OdsSaveOptions | How to set LibreOffice generator when saving as .fods using Aspose.Cells | Specify ODF 1.2 strict version in Aspose.Cells OdsSaveOptions C# | Save Excel workbook as Flat ODS (FODS) using Aspose.Cells API
// Tags: Aspose.Cells OdsSaveOptions flat ODS export | C# save workbook as FODS | set LibreOffice generator OdsSaveOptions | configure ODF strict version Aspose.Cells | export to Flat ODS using Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Creates a new Workbook, adds sample text and a date, configures OdsSaveOptions with the LibreOffice generator and ODF 1.2 strict version, and saves the workbook as a Flat ODS (.fods) file named output.fods.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Create OdsSaveOptions for saving as Flat ODS (FODS)
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        // Optional: specify the generator type
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;
        // Optional: set the ODF version
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

        // Save the workbook as FODS using the save options
        workbook.Save("output.fods", saveOptions);
    }
}
