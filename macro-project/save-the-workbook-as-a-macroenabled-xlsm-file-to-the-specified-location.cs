// Title: How to save a workbook as a macro‑enabled XLSM file with Aspose.Cells for .NET (C# example)
// Common Searches: C# Aspose.Cells save workbook as macro enabled XLSM file | Aspose.Cells SaveFormat.Xlsm usage in .NET application | Create and export macro‑enabled Excel workbook using Aspose.Cells C# | Specify output path when saving XLSM with Aspose.Cells | Generate XLSM file programmatically with Aspose.Cells library
// Tags: Aspose.Cells SaveFormat.Xlsm usage | C# macro-enabled Excel file generation | export workbook to XLSM with Aspose.Cells | define output path for XLSM save | programmatic creation of macro-enabled workbook

using Aspose.Cells;
using System;

// The sample creates a new Workbook, adds data to the first worksheet, sets a file path, and saves the workbook as a macro‑enabled XLSM file using Aspose.Cells' SaveFormat.Xlsm in C#.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Example: add some data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Text");
        sheet.Cells["B1"].PutValue(12345);

        // Specify the full path where the XLSM file will be saved
        string outputPath = @"C:\Temp\MyMacroEnabledWorkbook.xlsm";

        // Save the workbook as a macro‑enabled XLSM file
        workbook.Save(outputPath, SaveFormat.Xlsm);
    }
}
