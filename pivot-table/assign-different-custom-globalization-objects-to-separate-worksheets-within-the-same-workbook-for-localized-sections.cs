// Title: How to assign different custom globalization objects to separate worksheets in an Aspose.Cells workbook (C#)
// AI Prompts: Create two classes inheriting SettableGlobalizationSettings (e.g., EnglishGlobalization and GermanGlobalization) and configure boolean strings and function names. | Set the workbook's GlobalizationSettings to each class before adding a worksheet, then write localized values and formulas on the respective sheets. | Calculate all formulas, output the localized boolean and formula results to the console, and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# set different globalization settings for each worksheet | per‑sheet localization of boolean strings and function names in Excel using Aspose.Cells | example of English and German custom globalization in a single workbook Aspose.Cells | how to apply SettableGlobalizationSettings to individual sheets in Aspose.Cells | save workbook with mixed language globalization Aspose.Cells C#
// Tags: per‑sheet custom globalization Aspose.Cells | localized function names worksheet C# | boolean string localization Aspose.Cells | SettableGlobalizationSettings example C# | mixed language workbook Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // First custom globalization: English‑like settings
    // The example defines EnglishGlobalization and GermanGlobalization classes derived from SettableGlobalizationSettings, assigns each to a different worksheet in the same Workbook, writes boolean values and localized formulas, calculates the formulas, prints the localized results to the console, and saves the file as WorkbookPerSheetGlobalization.xlsx.
    public class EnglishGlobalization : SettableGlobalizationSettings
    {
        public EnglishGlobalization()
        {
            // Boolean strings
            SetBooleanValueString(true, "TRUE_EN");
            SetBooleanValueString(false, "FALSE_EN");
            // Function names
            SetLocalFunctionName("SUM", "SUM_EN", true);
        }
    }

    // Second custom globalization: German‑like settings
    public class GermanGlobalization : SettableGlobalizationSettings
    {
        public GermanGlobalization()
        {
            SetBooleanValueString(true, "WAHR");
            SetBooleanValueString(false, "FALSCH");
            SetLocalFunctionName("SUM", "SUMME", true);
        }
    }

    public class WorkbookWithPerSheetGlobalization
    {
        public static void Run()
        {
            // ---------- Create a new workbook ----------
            Workbook wb = new Workbook();

            // ---------- Apply English globalization and add first sheet ----------
            wb.Settings.GlobalizationSettings = new EnglishGlobalization();

            // Adding a sheet uses the current globalization settings for the default name
            int firstIndex = wb.Worksheets.Add();               // Sheet name will be based on English settings
            Worksheet sheetEn = wb.Worksheets[firstIndex];
            sheetEn.Name = "EnglishSection";                    // Optional explicit rename
            sheetEn.Cells["A1"].PutValue(true);                // Will display "TRUE_EN"
            sheetEn.Cells["A2"].Formula = "=SUM(B1:B3)";       // Uses "SUM_EN" internally

            // ---------- Apply German globalization and add second sheet ----------
            wb.Settings.GlobalizationSettings = new GermanGlobalization();

            int secondIndex = wb.Worksheets.Add();              // Sheet name will be based on German settings
            Worksheet sheetDe = wb.Worksheets[secondIndex];
            sheetDe.Name = "GermanSection";
            sheetDe.Cells["A1"].PutValue(true);                // Will display "WAHR"
            sheetDe.Cells["A2"].Formula = "=SUMME(B1:B3)";     // Localized function name works

            // ---------- Populate some data for the formulas ----------
            sheetEn.Cells["B1"].PutValue(10);
            sheetEn.Cells["B2"].PutValue(20);
            sheetEn.Cells["B3"].PutValue(30);

            sheetDe.Cells["B1"].PutValue(5);
            sheetDe.Cells["B2"].PutValue(15);
            sheetDe.Cells["B3"].PutValue(25);

            // Calculate formulas for both sheets
            wb.CalculateFormula();

            // ---------- Display results in console ----------
            Console.WriteLine($"English sheet boolean display: {sheetEn.Cells["A1"].StringValue}");
            Console.WriteLine($"English sheet SUM result: {sheetEn.Cells["A2"].Value}");

            Console.WriteLine($"German sheet boolean display: {sheetDe.Cells["A1"].StringValue}");
            Console.WriteLine($"German sheet SUM result: {sheetDe.Cells["A2"].Value}");

            // ---------- Save the workbook ----------
            wb.Save("WorkbookPerSheetGlobalization.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            WorkbookWithPerSheetGlobalization.Run();
        }
    }
}
