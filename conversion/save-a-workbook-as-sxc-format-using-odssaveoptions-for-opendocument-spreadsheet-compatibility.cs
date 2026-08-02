// Title: C# – Save Workbook as StarOffice Calc (.sxc) with Aspose.Cells OdsSaveOptions
// Description: Shows how to create a workbook, configure OdsSaveOptions with SaveFormat.Sxc, optionally set the LibreOffice generator type, and save the result as a StarOffice Calc (.sxc) OpenDocument Spreadsheet using Aspose.Cells.
// Keywords: Aspose.Cells | C# | OdsSaveOptions | SaveFormat.Sxc | StarOffice Calc | SXC export | OpenDocument Spreadsheet | LibreOffice generator | Excel to SXC conversion | C# workbook save as sxc
// Common Searches: Aspose.Cells save as .sxc C# | How to export Excel to StarOffice Calc using Aspose | OdsSaveOptions Sxc format example | C# convert .xlsx to .sxc with Aspose | Set generator type LibreOffice when saving SXC | OpenDocument spreadsheet conversion .NET
// Developer Intent: Generate a StarOffice Calc (.sxc) file from a .NET workbook via Aspose.Cells.
// Use Cases: Create cross‑platform .sxc reports directly from a C# application. | Produce LibreOffice‑compatible spreadsheets by specifying the LibreOffice generator type. | Automate bulk conversion of existing Excel files to the SXC OpenDocument format.
// AI Prompts: Write C# code that converts an existing .xlsx file to .sxc using Aspose.Cells and sets the LibreOffice generator type. | Provide a reusable method that accepts a Workbook object and optional OdsSaveOptions, then saves it as .sxc. | Explain which OdsSaveOptions properties affect compatibility when exporting to the SXC format.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Shows how to create a workbook, configure OdsSaveOptions with SaveFormat.Sxc, optionally set the LibreOffice generator type, and save the result as a StarOffice Calc (.sxc) OpenDocument Spreadsheet using Aspose.Cells.
class SaveAsSxcDemo
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Hello SXC");

        // Create ODS save options specifying the SXC format
        OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.Sxc);
        // Optional: set the generator type for better compatibility
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Save the workbook as StarOffice Calc Spreadsheet (.sxc)
        workbook.Save("output.sxc", saveOptions);
    }
}
