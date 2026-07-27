// Title: C# – Open ODS workbook, set ODF 1.1 strict version, and save with Aspose.Cells
// Description: Demonstrates loading an ODS file with Aspose.Cells, configuring OdsSaveOptions.OdfStrictVersion to Odf11, and saving the workbook as a fully ODF 1.1‑strict compliant document.
// Keywords: Aspose.Cells ODS strict version | C# OdfStrictVersion Odf11 | OdsSaveOptions example | save ODS as ODF 1.1 | OpenDocument Format 1.1 compliance | LibreOffice ODS compatibility | Aspose.Cells .NET ODS conversion
// Common Searches: Aspose.Cells set ODF 1.1 strict version C# | How to save ODS with ODF 1.1 compliance using Aspose.Cells | OdsSaveOptions OdfStrictVersion example | Convert ODS to ODF 1.1 strict with .NET | C# code for ODS strict version Aspose
// Developer Intent: Open an existing ODS workbook, enforce ODF 1.1 strict compliance, and save the result using Aspose.Cells for .NET.
// Use Cases: Create ODS reports that must pass ODF 1.1 strict validation for LibreOffice or other ODF‑compatible tools. | Upgrade legacy ODS files to the latest strict specification before distribution or archival. | Automate batch processing of ODS documents, ensuring every output adheres to ODF 1.1 strict standards.
// AI Prompts: Generate C# code with Aspose.Cells that loads an ODS file, sets OdfStrictVersion to Odf11, and saves it as ODF 1.1 strict. | Explain how OdsSaveOptions.OdfStrictVersion influences the generated ODS file and list all OpenDocumentFormatVersionType values. | Provide a C# loop that converts multiple ODS files to ODF 1.1 strict format using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsOdsStrictVersionDemo
{
    // Demonstrates loading an ODS file with Aspose.Cells, configuring OdsSaveOptions.OdfStrictVersion to Odf11, and saving the workbook as a fully ODF 1.1‑strict compliant document.
    class Program
    {
        static void Main()
        {
            // Load an existing ODS workbook
            Workbook workbook = new Workbook("input.ods");

            // Create ODS save options and set ODF version to 1.1 (strict)
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf11;

            // Save the workbook with ODF 1.1 compliance
            workbook.Save("output_strict.ods", saveOptions);
        }
    }
}
