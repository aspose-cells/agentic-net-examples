// Title: Export a Workbook to Flat ODS (FODS) with Aspose.Cells OdsSaveOptions in C#
// Description: Shows how to create a workbook, fill cells, configure OdsSaveOptions for the Flat ODS format (SaveFormat.Fods), set the LibreOffice generator type, and save the result as a .fods file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | OdsSaveOptions | Flat ODS | FODS | SaveFormat.Fods | LibreOffice generator | Excel to ODS conversion | .NET spreadsheet export | ODS XML
// Common Searches: Aspose.Cells save as FODS C# | How to use OdsSaveOptions for Flat ODS | Convert Excel to .fods with .NET | Set generator type LibreOffice Aspose.Cells | Flat ODS export example C#
// Developer Intent: Generate a Flat ODS (FODS) file from an Excel workbook using Aspose.Cells in a C# application.
// Use Cases: Create a lightweight XML representation of a spreadsheet for compatibility with LibreOffice or other ODS tools. | Embed generator metadata (e.g., LibreOffice) to identify the source application in the Flat ODS output. | Automate batch conversion of .xlsx files to .fods in server‑side .NET processes. | Preserve text, numeric, and date values in a flat XML format suitable for version control.
// AI Prompts: Write C# code that reads an existing .xlsx file, applies OdsSaveOptions with a custom generator, and saves it as a .fods file using Aspose.Cells. | Explain how to enable compression or add custom namespaces when saving Flat ODS with Aspose.Cells. | Show a .NET console application that iterates through a directory and converts all Excel workbooks to FODS using OdsSaveOptions. | Describe the differences between ODS and FODS and advise when to choose each format in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Shows how to create a workbook, fill cells, configure OdsSaveOptions for the Flat ODS format (SaveFormat.Fods), set the LibreOffice generator type, and save the result as a .fods file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(DateTime.Now);

        // Create OdsSaveOptions for the Flat ODS (FODS) format
        OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.Fods);
        // Optional: specify the generator type (LibreOffice in this example)
        saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;

        // Save the workbook as a .fods file using the specified options
        workbook.Save("output.fods", saveOptions);
    }
}
