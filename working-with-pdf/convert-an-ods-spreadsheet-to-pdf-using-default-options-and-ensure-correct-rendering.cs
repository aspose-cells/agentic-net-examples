// Title: Convert an ODS spreadsheet to PDF with Aspose.Cells for .NET using default options
// AI Prompts: Write C# code that loads an .ods workbook using Aspose.Cells and saves it as a PDF with the library's default save settings. | Generate a minimal Aspose.Cells .NET snippet to convert an OpenDocument Spreadsheet file to PDF without customizing any conversion options.
// Common Searches: Aspose.Cells C# convert ODS file to PDF with default settings | sample code to save OpenDocument spreadsheet as PDF using Aspose.Cells .NET | how to export .ods workbook to PDF in C# without specifying save options | default PDF rendering of ODS in Aspose.Cells example
// Tags: Aspose.Cells ODS to PDF conversion | Workbook.Save default PDF export | C# OpenDocument spreadsheet PDF rendering | Aspose.Cells .NET default PDF options | Aspose.Cells ODS workbook PDF export

using System;
using Aspose.Cells;

// The example loads an ODS file (input.ods) into an Aspose.Cells Workbook and saves it as a PDF (output.pdf) using the library's default PDF save options, ensuring proper rendering.
class Program
{
    static void Main()
    {
        // Load the ODS spreadsheet from file using default load options
        Workbook workbook = new Workbook("input.ods");

        // Save the workbook as PDF using default save options (ensures correct rendering)
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
