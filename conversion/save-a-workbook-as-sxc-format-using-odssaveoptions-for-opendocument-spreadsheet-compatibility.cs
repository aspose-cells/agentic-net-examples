// Title: Save a workbook as StarOffice Calc (.sxc) with Aspose.Cells using OdsSaveOptions in C#
// AI Prompts: Write C# code that creates a new Workbook, configures OdsSaveOptions with SaveFormat.Sxc and LibreOffice generator, and writes the file to disk as a .sxc document. | Demonstrate how to export an Aspose.Cells workbook to the StarOffice Calc format by setting the appropriate OdsSaveOptions in a .NET application.
// Common Searches: aspocells c# export workbook to sxc file | odsSaveOptions SaveFormat.Sxc example code | how to set LibreOffice generator for sxc output using Aspose.Cells | convert Excel workbook to StarOffice Calc (.sxc) in .NET | saving spreadsheet as OpenDocument SXC with Aspose.Cells
// Tags: Aspose.Cells OdsSaveOptions SaveFormat.Sxc | C# generate SXC file with Aspose.Cells | LibreOffice generator type for SXC output | OpenDocument spreadsheet compatibility Aspose | convert Excel to .sxc using Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// // This example creates a new Workbook, adds sample data, configures OdsSaveOptions with SaveFormat.Sxc and LibreOffice generator type, and saves the workbook as output.sxc.
class SaveAsSxc
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Create OdsSaveOptions specifying the SXC format
        OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.Sxc);
        // Set the generator type for better OpenDocument compatibility
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Save the workbook as StarOffice Calc Spreadsheet (.sxc)
        workbook.Save("output.sxc", saveOptions);
    }
}
