// Title: How to clear the Ribbon UI and export a workbook to ODS with Aspose.Cells for .NET
// AI Prompts: Generate C# code that sets Workbook.RibbonXml to an empty string and saves the workbook as an ODS file using Aspose.Cells OdsSaveOptions. | Show an example of creating a new Workbook, disabling the Ribbon UI, and persisting the changes when exporting to the OpenDocument Spreadsheet format with Aspose.Cells. | Write a snippet that demonstrates disabling the Ribbon UI before calling Workbook.Save with OdsSaveOptions in a .NET application.
// Common Searches: aspnet clear ribbonxml before saving workbook as .ods with Aspose.Cells | c# disable ribbon UI in Excel file and export to ODS using Aspose | how to persist ribbon removal when converting Excel to ODS in .NET | Aspose.Cells OdsSaveOptions save workbook without ribbon interface | remove Ribbon UI from workbook programmatically and save as ODS
// Tags: Workbook.RibbonXml empty string | ODS export with Aspose.Cells | disable ribbon UI programmatically | Aspose.Cells OdsSaveOptions usage | persist UI settings in ODS file

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Creates a new Workbook, clears its RibbonXml property to turn off the Ribbon UI, and saves the file as Result.ods using default OdsSaveOptions.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Disable the Ribbon UI by clearing the RibbonXml property
        workbook.RibbonXml = string.Empty;

        // Prepare ODS save options (default configuration)
        OdsSaveOptions odsOptions = new OdsSaveOptions();

        // Save the workbook as an ODS file with the specified options
        workbook.Save("Result.ods", odsOptions);
    }
}
