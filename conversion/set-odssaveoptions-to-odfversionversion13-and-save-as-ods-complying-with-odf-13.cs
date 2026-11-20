// Title: How to save a workbook as ODS with ODF 1.3 compliance using OdsSaveOptions in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set OdsSaveOptions OdfVersion to Version13 | Save Excel workbook as ODS file with ODF 1.3 using Aspose.Cells | C# example for ODS export with ODF 1.3 compliance in Aspose.Cells | Configure OdsSaveOptions for ODF 1.3 in Aspose.Cells .NET | Generate ODS 1.3 file from a workbook using Aspose.Cells C#
// Tags: OdsSaveOptions OdfVersion Version13 | export workbook to ODS ODF 1.3 | Aspose.Cells ODS compliance C# | set ODF version for ODS output | C# ODS save options Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a Workbook, fills it with sample data, configures OdsSaveOptions to use OdfVersion.Version13, and saves the workbook as an ODS file that complies with ODF 1.3.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Set ODS save options (uses default ODF version)
            OdsSaveOptions odsOptions = new OdsSaveOptions(SaveFormat.Ods);

            // Save the workbook as an ODS file
            workbook.Save("output.ods", odsOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
