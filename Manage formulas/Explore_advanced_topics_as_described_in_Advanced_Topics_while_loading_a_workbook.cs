using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Metadata;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "AdvancedTopicsDemo.xlsx";

        // If the source file does not exist, create a simple workbook for demonstration
        if (!File.Exists(sourcePath))
        {
            var tempWb = new Workbook();
            var tempWs = tempWb.Worksheets[0];
            tempWs.Name = "Data";
            tempWs.Cells["A1"].PutValue("Hello");
            tempWs.Cells["A1"].Formula = "=SUM(1,2,3)";
            tempWb.Save(sourcePath, SaveFormat.Xlsx);
        }

        // --------------------------------------------------------------------
        // Create LoadOptions with advanced settings
        // --------------------------------------------------------------------
        var loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            ParsingFormulaOnOpen = false,
            MemorySetting = MemorySetting.MemoryPreference,
            LoadFilter = new CustomLoadFilter()
        };

        // --------------------------------------------------------------------
        // Load the workbook using the constructor that accepts a file path and LoadOptions
        // --------------------------------------------------------------------
        var workbook = new Workbook(sourcePath, loadOptions);

        // --------------------------------------------------------------------
        // Access worksheet data to demonstrate that loading succeeded
        // --------------------------------------------------------------------
        Worksheet loadedWs = workbook.Worksheets[0];
        Console.WriteLine($"Worksheet Name: {loadedWs.Name}");
        Console.WriteLine($"Cell A1 Value: {loadedWs.Cells["A1"].Value}");
        Console.WriteLine($"Cell A1 Formula (if any): {loadedWs.Cells["A1"].Formula}");

        // --------------------------------------------------------------------
        // Extract built‑in document properties using WorkbookMetadata
        // --------------------------------------------------------------------
        var metaOptions = new MetadataOptions(MetadataType.DocumentProperties);
        var metadata = new WorkbookMetadata(sourcePath, metaOptions);
        Console.WriteLine("\nBuilt‑in Document Properties:");
        foreach (var prop in metadata.BuiltInDocumentProperties)
        {
            Console.WriteLine($"{prop.Name}: {prop.Value}");
        }

        // --------------------------------------------------------------------
        // Convert the loaded workbook to PDF using ConversionUtility
        // --------------------------------------------------------------------
        string pdfPath = "AdvancedTopicsDemo.pdf";
        ConversionUtility.Convert(sourcePath, pdfPath);
        Console.WriteLine($"\nWorkbook converted to PDF: {pdfPath}");

        // --------------------------------------------------------------------
        // Save the workbook (as loaded) to a new XLSX file
        // --------------------------------------------------------------------
        string savedPath = "AdvancedTopicsDemo_Loaded.xlsx";
        workbook.Save(savedPath, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved to: {savedPath}");
    }

    // ------------------------------------------------------------------------
    // Custom LoadFilter implementation to demonstrate selective loading
    // ------------------------------------------------------------------------
    private class CustomLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // For a sheet named "Data", load everything.
            // For all other sheets, load only cell values (no formulas, formatting, etc.).
            if (sheet.Name.Equals("Data", StringComparison.OrdinalIgnoreCase))
            {
                LoadDataFilterOptions = LoadDataFilterOptions.All;
            }
            else
            {
                LoadDataFilterOptions = LoadDataFilterOptions.CellValue;
            }
        }
    }
}