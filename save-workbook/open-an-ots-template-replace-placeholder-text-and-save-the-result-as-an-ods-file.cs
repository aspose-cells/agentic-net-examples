// Title: Load an OTS template, replace {Name} placeholder, and save as ODS using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an OTS file with Aspose.Cells, substitutes the token {Name} with a real value, and writes the workbook out as an ODS document. | Illustrate setting the OdsSaveOptions.GeneratorType property to LibreOffice while saving a modified workbook to ODS format. | Explain the steps to programmatically edit an OTS template and generate an ODS spreadsheet using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# replace token in OTS spreadsheet template | save modified OTS workbook as ODS with LibreOffice generator | example of OdsSaveOptions usage in .NET for ODS export | programmatically convert OTS to ODS using Aspose.Cells library
// Tags: replace token in OTS template Aspose.Cells | save workbook as ODS with OdsSaveOptions | C# ODS export using LibreOffice generator | load OTS file Aspose.Cells .NET | programmatic OTS to ODS conversion

using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

// The example loads an OTS template into a Workbook, replaces the {Name} placeholder with "John Doe", configures OdsSaveOptions to use the LibreOffice generator, and saves the updated workbook as an ODS file.
class Program
{
    static void Main()
    {
        // Path to the OTS template file
        string templatePath = "template.ots";

        // Load the OTS template into a Workbook instance
        Workbook workbook = new Workbook(templatePath);

        // Replace placeholder text (e.g., {Name}) with actual value
        workbook.Replace("{Name}", "John Doe");

        // Configure ODS save options (optional: set generator type)
        OdsSaveOptions saveOptions = new OdsSaveOptions
        {
            GeneratorType = OdsGeneratorType.LibreOffice
        };

        // Save the modified workbook as an ODS file
        string outputPath = "result.ods";
        workbook.Save(outputPath, saveOptions);
    }
}
