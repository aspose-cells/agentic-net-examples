// Title: Convert an ODS spreadsheet to PDF with Aspose.Cells for .NET using default options
// Common Searches: C# Aspose.Cells convert ODS file to PDF with default settings | How to export an OpenDocument spreadsheet as PDF using Aspose.Cells .NET | Aspose.Cells workbook.Save PDF example for .ods files | Default PDF export options for ODS in Aspose.Cells C# | Programmatically convert OpenDocument spreadsheet to PDF with Aspose.Cells
// Tags: ODS to PDF conversion Aspose.Cells .NET | Aspose.Cells default PDF save options | Workbook.Save PDF format C# | OpenDocument spreadsheet export Aspose.Cells | SaveFormat.Pdf usage in Aspose.Cells

using Aspose.Cells;
using System;

// Loads an ODS file into an Aspose.Cells Workbook and saves it as a PDF using the library's default PDF save options.
class Program
{
    static void Main()
    {
        // Load the ODS spreadsheet from file
        Workbook workbook = new Workbook("input.ods");

        // Convert and save the workbook to PDF using default settings
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
