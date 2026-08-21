// Title: Save Excel as ODS (ODF 1.3) with Aspose.Cells – C# Example
// Description: Demonstrates how to create a workbook, add data, configure OdsSaveOptions with OdfStrictVersion = Odf13, and export the file as an ODS document that fully complies with the ODF 1.3 specification using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | ODS export | ODF 1.3 | OdsSaveOptions | OdfStrictVersion | Odf13 | Excel to ODS conversion | OpenDocument format | LibreOffice compatibility
// Common Searches: Aspose.Cells set ODF version 1.3 C# | OdsSaveOptions OdfStrictVersion example | Export Excel to ODS strict ODF 1.3 | Save workbook as ODS using Aspose.Cells .NET | How to create ODS 1.3 file with Aspose.Cells
// Developer Intent: Configure OdsSaveOptions for ODF 1.3 strict mode and save a workbook as an ODS file.
// Use Cases: Produce ODS reports that must pass ODF 1.3 validation for LibreOffice or OpenOffice users. | Automate batch conversion of Excel files to ODS 1.3 in a .NET backend service. | Generate ODS documents with a guaranteed version for regulatory or archival purposes.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as ODS using OdfStrictVersion = Odf13. | Explain how OdsSaveOptions.OdfStrictVersion maps to ODF versions and how to verify ODS compliance. | Provide a step‑by‑step guide to convert an existing .xlsx file to an ODS file that meets ODF 1.3 standards using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add data, configure OdsSaveOptions with OdfStrictVersion = Odf13, and export the file as an ODS document that fully complies with the ODF 1.3 specification using Aspose.Cells for .NET.
    public class OdsSaveWithVersion13Demo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and add some sample data
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("ODF 1.3 Demo");
                worksheet.Cells["A2"].PutValue(123);
                worksheet.Cells["A3"].PutValue(DateTime.Now);

                // Create ODS save options and set ODF version to 1.3 (strict)
                OdsSaveOptions saveOptions = new OdsSaveOptions
                {
                    OdfStrictVersion = OpenDocumentFormatVersionType.Odf13
                };

                // Save the workbook as ODS with the specified ODF version
                workbook.Save("OdfVersion13Demo.ods", saveOptions);
                Console.WriteLine("Workbook saved successfully as OdfVersion13Demo.ods");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            OdsSaveWithVersion13Demo.Run();
        }
    }
}
