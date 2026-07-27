// Title: Save Excel to ODS with ODF 1.2 compliance using Aspose.Cells for .NET
// Description: Shows how to create a Workbook, add sample data, set OdsSaveOptions.OdfStrictVersion to Odf12, and save the file as an ODS document that conforms to the ODF 1.2 specification.
// Keywords: Aspose.Cells | C# ODS export | ODF 1.2 | OdsSaveOptions | OpenDocumentFormatVersionType | Excel to ODS conversion | .NET spreadsheet library | strict ODF version
// Common Searches: Aspose.Cells save as ODS ODF 1.2 | C# OdsSaveOptions OdfStrictVersion example | Export Excel to ODS with strict ODF 1.2 | How to set ODF version when saving ODS in .NET | Aspose.Cells ODS strict compliance code
// Developer Intent: Generate an ODS file from a .NET workbook that strictly follows the ODF 1.2 standard.
// Use Cases: Produce ODS reports compatible with LibreOffice and other ODF 1.2 readers. | Exchange data with partners that require ODF 1.2‑compliant documents. | Automate batch conversion of Excel files to ODS while enforcing strict ODF versioning for archival or regulatory purposes.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as ODS with ODF 1.2 and apply custom cell formatting. | Explain the OdsSaveOptions properties for ODF version control and list other ODS export options available in Aspose.Cells. | Create a script that processes multiple Excel files and saves each as an ODS document using the Odf12 strict version.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Shows how to create a Workbook, add sample data, set OdsSaveOptions.OdfStrictVersion to Odf12, and save the file as an ODS document that conforms to the ODF 1.2 specification.
class SaveOdsWithOdf12
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B1"].PutValue(DateTime.Now);

        // Create ODS save options and set the ODF version to 1.2
        OdsSaveOptions saveOptions = new OdsSaveOptions();
        saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

        // Save the workbook as ODS using the specified options
        workbook.Save("SampleOdf12.ods", saveOptions);
    }
}
