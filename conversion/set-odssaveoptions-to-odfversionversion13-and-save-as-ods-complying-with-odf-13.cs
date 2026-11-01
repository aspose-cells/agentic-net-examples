// Title: How to save a workbook as ODS with ODF 1.3 compliance using Aspose.Cells for .NET
// Common Searches: Aspose.Cells set OdsSaveOptions to ODF 1.3 version in C# | Save Excel workbook as ODS compliant with ODF 1.3 using Aspose.Cells .NET | C# code to export a workbook to ODS with latest ODF version in Aspose.Cells | Why does OdsSaveOptions not expose an OdfVersion property in Aspose.Cells | Configure ODS save options for ODF 1.3 in Aspose.Cells for .NET
// Tags: OdsSaveOptions ODF 1.3 Aspose.Cells | export workbook to ODS .NET | Aspose.Cells ODS save options configuration | C# ODS compliance Aspose.Cells | SaveFormat.Ods usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a Workbook, initializing OdsSaveOptions with SaveFormat.Ods, noting that the API defaults to the latest ODF version (no explicit OdfVersion property), and saving the file as an ODS document compliant with ODF 1.3.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B1"].PutValue(12345);

            // Configure ODS save options (use the non‑obsolete SaveFormat)
            OdsSaveOptions odsOptions = new OdsSaveOptions(SaveFormat.Ods);
            // Current API does not expose a Version property; it defaults to the latest ODF version.

            // Save the workbook as an ODS file
            string outputPath = "output.ods";
            workbook.Save(outputPath, odsOptions);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
