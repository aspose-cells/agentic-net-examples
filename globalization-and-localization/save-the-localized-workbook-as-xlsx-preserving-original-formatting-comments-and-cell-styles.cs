// Title: Save a localized Excel workbook as XLSX with full formatting, comments, and styles using Aspose.Cells for .NET
// Description: Load an existing localized workbook (input.xlsx) with Aspose.Cells, then save it as output.xlsx in XLSX format while preserving all original formatting, cell comments, and style definitions. The example also demonstrates proper disposal of the Workbook object.
// Keywords: Aspose.Cells save XLSX | preserve Excel formatting .NET | retain cell comments Aspose | keep cell styles when saving | localized workbook conversion | SaveFormat.Xlsx example | Excel globalization Aspose.Cells | C# workbook dispose
// Common Searches: Aspose.Cells save workbook without losing formatting | How to keep comments when converting Excel to XLSX in .NET | Preserve cell styles during Excel localization export | Save localized Excel file as XLSX using C# | Aspose.Cells keep original design after Save
// Developer Intent: The developer needs to export a loaded localized Excel file to XLSX while ensuring that formatting, comments, and style information remain unchanged.
// Use Cases: Export a language‑specific Excel template to a new XLSX file after populating data, keeping the template’s visual layout intact. | Allow users of a web‑based reporting tool to download their edited workbook with all colors, borders, and notes preserved. | Automate batch conversion of multiple localized workbooks to XLSX for archival, guaranteeing that every workbook retains its original design elements.
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, modifies a cell, and saves it as XLSX while preserving formatting, comments, and styles. | Show how to configure Aspose.Cells SaveOptions to ensure that all cell comments and style definitions are retained when saving to XLSX. | Explain best practices for disposing Aspose.Cells Workbook objects after saving to prevent memory leaks in a high‑throughput application.

using System;
using Aspose.Cells;

// Load an existing localized workbook (input.xlsx) with Aspose.Cells, then save it as output.xlsx in XLSX format while preserving all original formatting, cell comments, and style definitions. The example also demonstrates proper disposal of the Workbook object.
class Program
{
    static void Main()
    {
        // Load the existing localized workbook.
        // The constructor with a file path loads the workbook preserving all original content.
        Workbook workbook = new Workbook("input.xlsx");

        // Save the workbook as XLSX.
        // Using the Save(string, SaveFormat) overload ensures that all formatting,
        // comments, and cell styles are retained in the output file.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);

        // Clean up resources.
        workbook.Dispose();
    }
}
