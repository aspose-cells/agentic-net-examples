// Title: Convert a large XLSX workbook to PDF in C# using Aspose.Cells LightCells API with MemoryPreference
// Common Searches: how to convert a large xlsx file to pdf in C# using Aspose.Cells LightCells API | Aspose.Cells load large workbook with MemoryPreference to reduce memory usage | C# example of saving an Excel workbook directly to PDF with LoadOptions | using LoadOptions.MemorySetting.MemoryPreference for PDF export of big Excel files | lightcells API memory‑efficient Excel to PDF conversion in .NET
// Tags: lightcells api memorypreference loading | c# aspose.cells loadoptions memorysetting | excel to pdf conversion aspose.cells | large workbook pdf export c# | load large xlsx with lightcells

using System;
using Aspose.Cells;

// The sample demonstrates loading a large XLSX workbook with Aspose.Cells using LoadOptions set to MemorySetting.MemoryPreference (LightCells mode) to minimize memory consumption, then saving the workbook directly to PDF in C#.
class Program
{
    static void Main()
    {
        // Path to the large Excel workbook
        string inputPath = "largeWorkbook.xlsx";

        // Path where the PDF will be saved
        string outputPath = "largeWorkbook.pdf";

        // Configure load options to use memory‑efficient settings (LightCells style)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        // Prefer memory saving mode for large files
        loadOptions.MemorySetting = MemorySetting.MemoryPreference;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Save the workbook directly to PDF format
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}
