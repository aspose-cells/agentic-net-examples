// Title: C# – Save Workbook as StarOffice Calc (.sxc) with Aspose.Cells OdsSaveOptions
// Description: Shows how to create a workbook, add sample data, configure OdsSaveOptions for SaveFormat.Sxc (StarOffice Calc), optionally set the LibreOffice generator, and save the file as .sxc for OpenDocument Spreadsheet compatibility.
// Keywords: Aspose.Cells | C# | Save as SXC | OdsSaveOptions | SaveFormat.Sxc | StarOffice Calc | OpenDocument Spreadsheet | LibreOffice generator | export to .sxc | spreadsheet conversion C#
// Common Searches: Aspose.Cells save as sxc c# | How to export workbook to .sxc using OdsSaveOptions | C# convert Excel to StarOffice Calc format | Set LibreOffice generator type when saving as SXC | OpenDocument spreadsheet conversion with Aspose.Cells
// Developer Intent: Generate a .sxc file from a workbook using Aspose.Cells OdsSaveOptions in C#.
// Use Cases: Provide .sxc files for users of LibreOffice or Apache OpenOffice. | Create legacy StarOffice Calc reports from .NET applications. | Batch‑export Excel workbooks to OpenDocument Spreadsheet format with specific generator settings.
// AI Prompts: Write C# code that loads an existing .xlsx file and saves it as .sxc using Aspose.Cells OdsSaveOptions. | Explain how to customize fonts, styles, and other OdsSaveOptions properties when exporting to SXC. | Show a C# example that converts multiple workbooks to .sxc files with progress reporting.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Shows how to create a workbook, add sample data, configure OdsSaveOptions for SaveFormat.Sxc (StarOffice Calc), optionally set the LibreOffice generator, and save the file as .sxc for OpenDocument Spreadsheet compatibility.
class SaveAsSxcDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Create OdsSaveOptions specifying the SXC format
        OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.Sxc);
        // Optional: set the generator type for better compatibility
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Save the workbook as StarOffice Calc Spreadsheet (.sxc)
        workbook.Save("output.sxc", saveOptions);
    }
}
