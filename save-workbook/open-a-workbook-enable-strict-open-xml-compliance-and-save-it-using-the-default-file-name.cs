// Title: Save a new Aspose.Cells workbook with ISO/IEC 29500:2008 Strict OpenXML compliance using the default filename in C#
// AI Prompts: Instantiate a Workbook, assign OoxmlCompliance.Iso29500_2008_Strict to its Settings.Compliance, and save it as 'StrictComplianceWorkbook.xlsx' with Aspose.Cells for .NET. | Load an existing Excel file, enforce ISO 29500 strict OpenXML mode via the Settings.Compliance property, and persist the workbook using the default file name in a C# application.
// Common Searches: c# aspose.cells enable ISO 29500 strict OpenXML mode when saving workbook | how to set OoxmlCompliance to Iso29500_2008_Strict in Aspose.Cells | save workbook with default name after applying strict OpenXML compliance in .NET | Aspose.Cells Workbook.Settings.Compliance property usage example | default filename for saving Excel workbook with Aspose.Cells .NET
// Tags: OoxmlCompliance Iso29500_2008_Strict configuration | save workbook default filename Aspose.Cells | strict OpenXML mode C# Aspose.Cells | Workbook.Settings.Compliance property example | ISO 29500 strict compliance Excel export

using System;
using Aspose.Cells;

// // Creates a workbook, sets ISO/IEC 29500:2008 strict OpenXML compliance, and saves it as StrictComplianceWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Enable strict Open XML compliance (ISO/IEC 29500:2008 Strict)
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Save the workbook using a default file name
        string defaultFileName = "StrictComplianceWorkbook.xlsx";
        workbook.Save(defaultFileName);
    }
}
