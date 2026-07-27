// Title: Save Excel to ODS with ODF 1.3 compliance using Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add sample data, set OdsSaveOptions.OdfStrictVersion to OpenDocumentFormatVersionType.Odf13, and save the file as an ODS document that conforms to the ODF 1.3 specification.
// Keywords: Aspose.Cells | C# | ODS export | ODF 1.3 | OdsSaveOptions | OpenDocumentFormatVersionType | Excel to ODS conversion | strict ODF | LibreOffice compatibility
// Common Searches: Aspose.Cells export to ODS ODF 1.3 | C# OdsSaveOptions OdfStrictVersion example | Create ODF 1.3 compliant ODS with Aspose | Save workbook as ODS using Aspose.Cells .NET
// Developer Intent: Export a workbook to an ODS file that meets the ODF 1.3 standard.
// Use Cases: Generate ODS reports that must pass ODF 1.3 validation for enterprise document workflows. | Archive Excel data in a format compatible with LibreOffice, OpenOffice, and other strict OpenDocument suites. | Automate bulk conversion of Excel files to ODS 1.3 for integration with third‑party ODF‑based systems.
// AI Prompts: Write C# code with Aspose.Cells that saves a workbook as ODS while setting OdfStrictVersion to Odf13 and applying custom cell styles. | Explain methods to validate an ODS file against the ODF 1.3 schema after it is generated with Aspose.Cells. | Show how to enable compression and other OdsSaveOptions features without breaking ODF 1.3 compliance.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsOdsVersionDemo
{
    // Shows how to create a workbook, add sample data, set OdsSaveOptions.OdfStrictVersion to OpenDocumentFormatVersionType.Odf13, and save the file as an ODS document that conforms to the ODF 1.3 specification.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("ODF 1.3 Demo");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B1"].PutValue(DateTime.Now);

            // Create ODS save options
            OdsSaveOptions saveOptions = new OdsSaveOptions();

            // Set the ODF strict version to 1.3
            saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf13;

            // Save the workbook as ODS complying with ODF 1.3
            workbook.Save("DemoOdf13.ods", saveOptions);
        }
    }
}
