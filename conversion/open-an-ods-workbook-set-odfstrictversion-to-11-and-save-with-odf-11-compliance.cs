// Title: Save ODS with ODF 1.1 Strict Compliance using Aspose.Cells for .NET
// Description: Load an existing ODS workbook, set the OdfStrictVersion to Odf11 via OdsSaveOptions, and save the file to meet ODF 1.1 strict specifications with Aspose.Cells for .NET.
// Keywords: Aspose.Cells OdsSaveOptions | OdfStrictVersion Odf11 | ODF 1.1 strict compliance | C# save ODS as ODF 1.1 | convert Excel to ODS 1.1 | Aspose.Cells ODS export
// Common Searches: Aspose.Cells set OdfStrictVersion to Odf11 | save ODS file with ODF 1.1 strict mode C# | how to enforce ODF 1.1 compliance using Aspose.Cells | OdsSaveOptions OdfStrictVersion example | export workbook to ODS 1.1 strict with .NET
// Developer Intent: The developer needs to open an ODS workbook, enforce ODF 1.1 strict version, and write the result as a compliant ODS file.
// Use Cases: Produce ODS reports that must pass ODF 1.1 strict validation for enterprise document workflows. | Batch‑upgrade legacy ODS files to ODF 1.1 strict format before archiving or distribution. | Generate regulatory or government submissions where ODF 1.1 strict compliance is mandatory.
// AI Prompts: Generate C# code that loads an ODS file, sets OdfStrictVersion to Odf11, and saves it with Aspose.Cells. | Explain the differences between OdfStrictVersion values and their effect on the resulting ODS file. | Create a C# loop that processes a folder of Excel files, converting each to an ODS file with ODF 1.1 strict compliance.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Load an existing ODS workbook, set the OdfStrictVersion to Odf11 via OdsSaveOptions, and save the file to meet ODF 1.1 strict specifications with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing ODS workbook
        string inputPath = "input.ods";
        Workbook workbook = new Workbook(inputPath);

        // Create ODS save options and set the ODF strict version to 1.1
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf11;

        // Save the workbook with ODF 1.1 compliance
        string outputPath = "output_strict11.ods";
        workbook.Save(outputPath, saveOptions);
    }
}
