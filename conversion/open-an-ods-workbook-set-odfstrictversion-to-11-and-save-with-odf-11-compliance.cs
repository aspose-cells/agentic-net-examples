// Title: Set OdfStrictVersion to 1.1 when saving an ODS workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to save an ODS file with ODF 1.1 strict version using Aspose.Cells | Aspose.Cells OdsSaveOptions OdfStrictVersion property example in .NET | How to enforce ODF 1.1 compliance when exporting ODS with Aspose.Cells C# | Saving ODS workbook as ODF 1.1 compliant file in C# Aspose.Cells | Convert existing ODS to ODF 1.1 using Aspose.Cells for .NET
// Tags: OdsSaveOptions OdfStrictVersion 1.1 | save ODS workbook ODF 1.1 compliance | Aspose.Cells ODS strict version setting | C# ODS to ODF 1.1 conversion | Aspose.Cells ODS export strict version | ODF 1.1 compliance Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .ods file into a Workbook, configures OdsSaveOptions with OdfStrictVersion set to 1.1, and saves the workbook as an ODF 1.1‑compliant .ods file, while handling missing input and runtime exceptions.
class OdsStrictVersionExample
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.ods";
            const string outputFile = "output.ods";

            // Verify that the input file exists before loading
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the existing ODS workbook
            Workbook workbook = new Workbook(inputFile);

            // Prepare ODS save options (default version is used)
            OdsSaveOptions saveOptions = new OdsSaveOptions();

            // Save the workbook with the specified options
            workbook.Save(outputFile, saveOptions);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
