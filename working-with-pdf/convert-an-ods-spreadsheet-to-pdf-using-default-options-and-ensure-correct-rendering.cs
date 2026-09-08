// Title: Convert an OpenDocument Spreadsheet (ODS) to PDF with Aspose.Cells for .NET using default settings
// AI Prompts: Generate a C# snippet that loads an ODS file and saves it as a PDF using Aspose.Cells with all default options. | Show a minimal .NET example that reads input.ods and exports output.pdf without configuring any PDF save parameters. | Provide step‑by‑step C# code to convert an OpenDocument Spreadsheet to PDF using Aspose.Cells' default workbook.Save behavior.
// Common Searches: asp.net convert ods to pdf using aspose.cells default options | c# Aspose.Cells example for exporting OpenDocument Spreadsheet as PDF | how to save an ODS workbook as PDF in .NET without custom settings | default PDF export of ODS files with Aspose.Cells C# | Aspose.Cells load ODS and generate PDF using built‑in defaults
// Tags: Aspose.Cells ODS to PDF conversion | C# load OpenDocument Spreadsheet | Aspose.Cells default PDF save options | export ODS workbook as PDF .NET | Aspose.Cells workbook.Save PDF format

using Aspose.Cells;
using System;

// Loads an ODS workbook named input.ods and saves it as output.pdf using Aspose.Cells with the default PDF save settings.
class Program
{
    static void Main()
    {
        // Load the ODS spreadsheet from file using default load options
        Workbook workbook = new Workbook("input.ods");

        // Save the workbook as PDF using default PDF save options
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
