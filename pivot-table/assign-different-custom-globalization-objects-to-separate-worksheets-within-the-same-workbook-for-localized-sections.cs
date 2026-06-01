using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom globalization for English (default)
    public class EnglishGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "TRUE_EN" : "FALSE_EN";
        }

        public override string GetErrorValueString(string err)
        {
            // Keep default error strings
            return base.GetErrorValueString(err);
        }
    }

    // Custom globalization for Russian
    public class RussianGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string err)
        {
            // Example of localized error messages
            return err switch
            {
                "#DIV/0!" => "#ДЕЛ/0!",
                "#VALUE!" => "#ЗНАЧ!",
                _ => base.GetErrorValueString(err)
            };
        }
    }

    public class WorkbookWithPerWorksheetGlobalization
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();

            // -------------------------------------------------
            // Worksheet 1 – English globalization
            // -------------------------------------------------
            wb.Settings.GlobalizationSettings = new EnglishGlobalizationSettings(); // apply English settings

            Worksheet sheetEn = wb.Worksheets[0];
            sheetEn.Name = "EnglishSheet";

            // Populate some boolean values and a formula
            sheetEn.Cells["A1"].PutValue(true);
            sheetEn.Cells["A2"].PutValue(false);
            sheetEn.Cells["A3"].Formula = "=TRUE()";   // will display "TRUE_EN"
            sheetEn.Cells["A4"].Formula = "=FALSE()";  // will display "FALSE_EN"

            // -------------------------------------------------
            // Worksheet 2 – Russian globalization
            // -------------------------------------------------
            // Add a new worksheet for Russian section
            int russianIndex = wb.Worksheets.Add();
            Worksheet sheetRu = wb.Worksheets[russianIndex];
            sheetRu.Name = "RussianSheet";

            // Switch globalization settings to Russian before filling the sheet
            wb.Settings.GlobalizationSettings = new RussianGlobalizationSettings(); // apply Russian settings

            // Populate boolean values and a formula that will use Russian strings
            sheetRu.Cells["A1"].PutValue(true);
            sheetRu.Cells["A2"].PutValue(false);
            sheetRu.Cells["A3"].Formula = "=TRUE()";   // will display "ИСТИНА"
            sheetRu.Cells["A4"].Formula = "=FALSE()";  // will display "ЛОЖЬ"

            // Demonstrate localized error handling
            sheetRu.Cells["B1"].Formula = "=1/0"; // will produce #DIV/0! error, displayed as #ДЕЛ/0!

            // -------------------------------------------------
            // Calculate formulas (required to materialize the localized strings)
            // -------------------------------------------------
            wb.CalculateFormula();

            // -------------------------------------------------
            // Save the workbook (lifecycle rule: save)
            // -------------------------------------------------
            wb.Save("WorkbookPerWorksheetGlobalization.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            WorkbookWithPerWorksheetGlobalization.Run();
            Console.WriteLine("Workbook created with per‑worksheet globalization settings.");
        }
    }
}