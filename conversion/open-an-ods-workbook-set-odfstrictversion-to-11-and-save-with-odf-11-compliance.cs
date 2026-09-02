// Title: Open an ODS workbook, set ODF 1.1 strict version, and save it with Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .ods file, configure OdsSaveOptions with OdfStrictVersion = Odf11, and save the workbook using Aspose.Cells in C#. | Create OdsSaveOptions for ODF 1.1 strict compliance and apply it when exporting a loaded workbook with Aspose.Cells for .NET. | Demonstrate setting the OdfStrictVersion property to Odf11 before saving an ODS workbook in C#.
// Common Searches: Aspose.Cells C# set OdfStrictVersion to Odf11 when saving ODS | How to enforce ODF 1.1 strict compliance in an ODS file using Aspose.Cells for .NET | C# example for loading an ODS workbook and saving with ODF 1.1 strict version | OdsSaveOptions OdfStrictVersion property usage in Aspose.Cells .NET | Convert ODS to ODF 1.1 strict format with Aspose.Cells C# code
// Tags: Aspose.Cells OdsSaveOptions OdfStrictVersion | C# ODS strict ODF 1.1 export | OpenDocumentFormatVersionType Odf11 example | save ODS with ODF 1.1 compliance Aspose.Cells | load and save ODS workbook .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Loads an existing ODS workbook, sets the OdfStrictVersion to ODF 1.1 via OdsSaveOptions, and saves the file as a strict‑compliant ODS document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the source ODS workbook
        string inputPath = "input.ods";

        // Load the existing ODS workbook
        Workbook workbook = new Workbook(inputPath);

        // Create ODS save options
        OdsSaveOptions saveOptions = new OdsSaveOptions();

        // Set the ODF version to strict 1.1 compliance
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf11;

        // Save the workbook with the specified ODF 1.1 compliance
        string outputPath = "output_strict11.ods";
        workbook.Save(outputPath, saveOptions);
    }
}
