// Title: Save a Workbook as ODS with ODF 1.2 compliance using Aspose.Cells (C#)
// Description: Shows how to create a workbook, add data to cell A1, configure OdsSaveOptions.OdfStrictVersion to Odf12, and save the file as OutputVersion12.ods, guaranteeing ODF 1.2 strict compliance.
// Keywords: Aspose.Cells | C# | ODS export | ODF 1.2 | OdsSaveOptions | OdfStrictVersion | OpenDocumentFormatVersionType | convert Excel to ODS | save as ODS .NET | ODF compliance
// Common Searches: Aspose.Cells save workbook as ODS ODF 1.2 | C# OdsSaveOptions OdfStrictVersion example | How to enforce ODF 1.2 when exporting to ODS | Convert Excel to ODS with strict ODF version using Aspose | Save workbook as ODS 1.2 using Aspose.Cells C#
// Developer Intent: Generate an ODS file from a workbook while enforcing ODF 1.2 strict version.
// Use Cases: Produce ODS reports that must pass ODF 1.2 validation for LibreOffice or other ODF‑compatible suites. | Archive Excel data as ODS files with guaranteed version compliance for long‑term storage. | Automate batch conversion of multiple Excel workbooks to ODS 1.2 in a .NET service or scheduled task.
// AI Prompts: Write C# code with Aspose.Cells to convert an existing Excel file to ODS while setting OdfStrictVersion to Odf12. | Explain the impact of OdsSaveOptions.OdfStrictVersion on the generated ODS file and list all supported enum values. | Provide a C# example that saves a workbook as ODS with ODF 1.2, enabling compression and password protection.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Shows how to create a workbook, add data to cell A1, configure OdsSaveOptions.OdfStrictVersion to Odf12, and save the file as OutputVersion12.ods, guaranteeing ODF 1.2 strict compliance.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and add some sample data
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Sample data for ODS");

        // Create ODS save options and set the ODF version to 1.2
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

        // Save the workbook as ODS using the specified options
        workbook.Save("OutputVersion12.ods", saveOptions);
    }
}
