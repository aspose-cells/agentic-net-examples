// Title: Change the default font of a FODS workbook and save it as an ODS file with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .fods spreadsheet, updates the workbook's default style to Calibri 11, and exports it as an .ods file using the LibreOffice generator. | Generate a complete Aspose.Cells example that opens a FODS file, sets the default font for the entire workbook, and saves the result as an ODS document with custom OdsSaveOptions.
// Common Searches: how to set the default workbook font when converting a FODS file to ODS in C# | Aspose.Cells example for loading FODS and saving as ODS with specific font | C# change default style font in a spreadsheet loaded from FODS | use OdsSaveOptions to specify LibreOffice generator while saving ODS | convert FODS to ODS and apply Calibri font using Aspose.Cells for .NET
// Tags: load FODS workbook Aspose.Cells | modify workbook default style font | save workbook as ODS using OdsSaveOptions | LibreOffice generator OdsSaveOptions C# | Aspose.Cells default font customization

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Loads a .fods file, changes the workbook's default font to Calibri size 11, and saves the workbook as an .ods file using the LibreOffice generator via Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the FODS file with default load options
        OdsLoadOptions loadOptions = new OdsLoadOptions();
        Workbook workbook = new Workbook("input.fods", loadOptions);

        // Change the workbook's default font
        workbook.DefaultStyle.Font.Name = "Calibri";
        workbook.DefaultStyle.Font.Size = 11;

        // Prepare ODS save options (optional configuration)
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Save the workbook as an ODS file
        workbook.Save("output.ods", saveOptions);
    }
}
