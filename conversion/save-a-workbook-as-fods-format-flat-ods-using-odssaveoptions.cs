// Title: Save a Workbook as Flat ODS (FODS) with Aspose.Cells C# using OdsSaveOptions
// Description: Demonstrates how to create a workbook, set OdsSaveOptions (generator type and ODF strict version), and export the file to Flat ODS (.fods) with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Flat ODS | FODS | OdsSaveOptions | ODF 1.2 | LibreOffice generator | Excel to ODS conversion
// Common Searches: Aspose.Cells export to FODS C# | Save workbook as Flat ODS using OdsSaveOptions | Set ODF version when saving .fods with Aspose | C# code for generating .fods file | How to specify generator type for FODS output
// Developer Intent: Create a Flat ODS file from an Aspose.Cells workbook in C# while controlling generator and ODF version settings.
// Use Cases: Generate a lightweight .fods document for ODF‑compatible editors. | Ensure compliance with ODF 1.2 strict specifications by setting the appropriate version. | Customize the file metadata by selecting a LibreOffice generator.
// AI Prompts: Write C# code that uses Aspose.Cells to save a workbook as a Flat ODS file with LibreOffice as the generator and ODF 1.2 strict version. | Show how to configure OdsSaveOptions for FODS output and save an existing workbook to "output.fods". | Explain how to adjust OdsSaveOptions properties to meet specific ODF version requirements in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Demonstrates how to create a workbook, set OdsSaveOptions (generator type and ODF strict version), and export the file to Flat ODS (.fods) with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello FODS");
        sheet.Cells["B2"].PutValue(12345);

        // Create OdsSaveOptions for saving as Flat ODS (FODS)
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        // Optional: specify the generator and ODF version
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

        // Save the workbook in FODS format using the SaveOptions overload
        workbook.Save("output.fods", saveOptions);
    }
}
