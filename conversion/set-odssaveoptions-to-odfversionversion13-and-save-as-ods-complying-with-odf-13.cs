// Title: Export a Workbook to ODS with ODF 1.3 compliance using Aspose.Cells OdsSaveOptions in C#
// AI Prompts: Generate C# code that creates a Workbook, populates cells, configures OdsSaveOptions for ODF 1.3, and writes an .ods file. | Show how to set the OdfStrictVersion property to Odf13 before saving a workbook as ODS with Aspose.Cells. | Provide a complete example of producing an ODF 1.3‑compatible ODS document from a C# application using Aspose.Cells.
// Common Searches: asp.net set OdfStrictVersion Odf13 Aspose.Cells | c# create ODS file with ODF 1.3 using Aspose.Cells | how to enforce ODF 1.3 compliance when exporting Excel to ODS in .NET | Aspose.Cells OdsSaveOptions version 1.3 sample code | saving workbook as ODS with strict ODF version in C#
// Tags: Aspose.Cells OdsSaveOptions ODF 1.3 | C# export workbook to ODS format | set OdfStrictVersion OpenDocumentFormatVersionType | ODF 1.3 compliant ODS output | Aspose.Cells ODS version configuration .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsOdsVersionDemo
{
    // The example creates a new Workbook, adds sample data to cells A1‑A3, configures OdsSaveOptions with OdfStrictVersion set to Odf13 for ODF 1.3 compliance, and saves the file as Output_Odf13.ods.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Demo for ODF 1.3");
            sheet.Cells["A2"].PutValue(12345);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Create ODS save options
            OdsSaveOptions saveOptions = new OdsSaveOptions();

            // Set the ODF version to 1.3 (ODF 1.3)
            saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf13;

            // Save the workbook as ODS using the specified options
            workbook.Save("Output_Odf13.ods", saveOptions);
        }
    }
}
