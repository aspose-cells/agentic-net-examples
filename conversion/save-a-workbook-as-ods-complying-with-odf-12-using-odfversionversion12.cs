// Title: Save an Aspose.Cells workbook as ODS with ODF 1.2 compliance using OdfVersion.Version12 in C#
// Common Searches: Aspose.Cells C# export workbook to ODS with ODF 1.2 | How to set OdfVersion.Version12 when saving as ODS in Aspose.Cells | C# code example for OdsSaveOptions ODF 1.2 compliance | Save Excel file as OpenDocument Spreadsheet using Aspose.Cells .NET | Configure OdsSaveOptions for ODF 1.2 in a .NET application
// Tags: Aspose.Cells OdsSaveOptions ODF 1.2 | C# ODS export with Aspose.Cells | OpenDocument Spreadsheet save options .NET | OdfVersion.Version12 configuration in Aspose.Cells | export Excel to ODS using Aspose.Cells API

using System;
using Aspose.Cells;

// The example creates a new Workbook, adds sample data, configures OdsSaveOptions to use OdfVersion.Version12 (ODF 1.2), and saves the file as an OpenDocument Spreadsheet (output.ods) while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Add sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Configure ODS save options (default ODF version is 1.2)
            OdsSaveOptions saveOptions = new OdsSaveOptions();

            // Save the workbook as ODS
            workbook.Save("output.ods", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
