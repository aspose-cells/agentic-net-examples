// Title: Save a Workbook as ODS with ODF 1.2 compliance using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a Workbook, configures OdsSaveOptions.OdfStrictVersion to Odf12, and saves it as an .ods file. | Show how to export multiple worksheets to ODS while enforcing ODF 1.2 compliance with Aspose.Cells. | Demonstrate adjusting OdsSaveOptions to meet strict ODF 1.2 requirements before calling Workbook.Save.
// Common Searches: asp.net export Excel to ODS with ODF 1.2 using Aspose.Cells | C# OdsSaveOptions OdfStrictVersion Odf12 example | how to set ODF strict version when converting workbook to OpenDocument Spreadsheet in .NET | save workbook as OpenDocument Spreadsheet 1.2 with Aspose.Cells C#
// Tags: Aspose.Cells OdsSaveOptions OdfStrictVersion | C# export workbook to ODS ODF 1.2 | set ODF version for OpenDocument Spreadsheet output | Aspose.Cells ODS conversion strict version | OpenDocument Format 1.2 compliance .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Creates a new Workbook, sets OdsSaveOptions.OdfStrictVersion to Odf12, and saves the file as an ODS document that complies with ODF 1.2.
class SaveWorkbookAsOdsWithOdf12
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["A2"].PutValue(12345);

        // Configure ODS save options to use ODF version 1.2
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

        // Save the workbook as ODS with the specified ODF version
        workbook.Save("OutputOdf12.ods", saveOptions);
    }
}
