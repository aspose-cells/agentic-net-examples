using System;
using System.IO;
using Aspose.Cells;

class WorkbookMergePrerequisiteDemo
{
    public static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        // Paths to source workbooks and the destination merged workbook
        string sourcePath1 = "Source1.xlsx";
        string sourcePath2 = "Source2.xlsx";
        string mergedPath = "MergedOutput.xlsx";

        // Create sample source workbooks if they do not exist
        CreateSampleWorkbook(sourcePath1, "Data from Source 1");
        CreateSampleWorkbook(sourcePath2, "Data from Source 2");

        // Load the source workbooks
        Workbook wb1 = new Workbook(sourcePath1);
        Workbook wb2 = new Workbook(sourcePath2);

        // ----- Prerequisite checks -----

        // 2. Write‑protection check – merging is not allowed on write‑protected workbooks
        if (wb1.Settings.WriteProtection.IsWriteProtected || wb2.Settings.WriteProtection.IsWriteProtected)
        {
            Console.WriteLine("Error: One of the workbooks is write protected. Aborting merge.");
            return;
        }

        // 3. Compatibility check – optional, just inform the user
        bool compatibilityOk = wb1.Settings.CheckCompatibility && wb2.Settings.CheckCompatibility;
        if (!compatibilityOk)
        {
            Console.WriteLine("Info: Compatibility checking is disabled on at least one workbook.");
        }

        // 4. Excel restriction check – ensure large strings or invalid data won't cause exceptions later
        bool restrictionCheck = wb1.Settings.CheckExcelRestriction && wb2.Settings.CheckExcelRestriction;
        Console.WriteLine($"Excel restriction checking is {(restrictionCheck ? "enabled" : "disabled")}.");

        // ----- Merge operation -----
        // Combine wb2 into wb1
        wb1.Combine(wb2);

        // ----- Save with validation of merged areas -----
        XlsbSaveOptions saveOptions = new XlsbSaveOptions
        {
            ValidateMergedAreas = true, // validates merged cells before saving
            MergeAreas = true           // merges conditional formatting and validation areas
        };

        // Save the merged workbook
        wb1.Save(mergedPath, saveOptions);
        Console.WriteLine($"Merged workbook saved successfully to '{mergedPath}'.");
    }

    // Helper method to create a simple workbook with merged cells
    private static void CreateSampleWorkbook(string filePath, string cellValue)
    {
        if (File.Exists(filePath))
            return; // Skip creation if the file already exists

        Workbook wb = new Workbook();                     // create new workbook
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue(cellValue);              // put sample data
        ws.Cells.Merge(0, 0, 1, 2);                       // merge A1:B1 to test validation
        wb.Save(filePath, SaveFormat.Xlsx);              // save using Workbook.Save(string, SaveFormat)
    }
}