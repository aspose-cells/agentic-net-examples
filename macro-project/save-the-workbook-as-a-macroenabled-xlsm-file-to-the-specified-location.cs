// Title: How to save a new Aspose.Cells workbook as a macro‑enabled XLSM file in C#
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, writes values to cells, and saves it as an XLSM file at a specified location. | Adapt the Aspose.Cells example to accept a folder path and filename from the user and store the macro‑enabled workbook there. | Show how to use Aspose.Cells SaveFormat.Xlsm to export a workbook that can contain VBA macros.
// Common Searches: aspnet c# save workbook as macro enabled xlsm using aspose.cells saveformat.xlsm | example code to export Excel file with macros using Aspose.Cells .NET | how to specify output directory when saving XLSM with Aspose.Cells in C# | create and save macro‑enabled Excel workbook programmatically with Aspose.Cells
// Tags: Aspose.Cells SaveFormat.Xlsm usage | C# create macro-enabled Excel workbook | export workbook to macro-enabled format .NET | set output path for Aspose.Cells save | macro-enabled file generation with Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new Workbook, renames the first worksheet to "Data", writes "Hello" and "World" into cells A1 and B1, defines an output path, and saves the workbook as a macro‑enabled XLSM file using SaveFormat.Xlsm.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and set its name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Add sample data to the worksheet
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Define the path where the macro‑enabled file will be saved
        string outputPath = @"C:\Temp\MyMacroEnabledWorkbook.xlsm";

        // Save the workbook as an XLSM (macro‑enabled) file
        workbook.Save(outputPath, SaveFormat.Xlsm);
    }
}
