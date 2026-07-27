using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // 1. Configure globalization settings to localize function names
        // ------------------------------------------------------------
        // Map the standard English function name "SUM" to the French localized name "SOMME"
        // The 'true' flag makes the mapping bidirectional, so both names can be used interchangeably.
        SettableGlobalizationSettings globalizationSettings = new SettableGlobalizationSettings();
        globalizationSettings.SetLocalFunctionName("SUM", "SOMME", true);
        workbook.Settings.GlobalizationSettings = globalizationSettings;

        // ------------------------------------------------------------
        // 2. Populate sample data for the formula
        // ------------------------------------------------------------
        worksheet.Cells["B1"].PutValue(10);
        worksheet.Cells["B2"].PutValue(20);
        worksheet.Cells["B3"].PutValue(30);

        // ------------------------------------------------------------
        // 3. Use the localized function name in a cell formula
        // ------------------------------------------------------------
        // The formula uses "SOMME", which will be resolved to the standard "SUM" during calculation.
        worksheet.Cells["A1"].Formula = "=SOMME(B1:B3)";

        // ------------------------------------------------------------
        // 4. Add a comment that contains text in its original language (Japanese)
        // ------------------------------------------------------------
        // The comment text is left untouched by globalization settings, preserving its language.
        int commentIdx = worksheet.Comments.Add("C5");
        Comment comment = worksheet.Comments[commentIdx];
        comment.Note = "これはテストコメントです"; // "This is a test comment" in Japanese
        comment.IsVisible = true;

        // ------------------------------------------------------------
        // 5. Calculate formulas and save the workbook
        // ------------------------------------------------------------
        workbook.CalculateFormula();
        workbook.Save("LocalizedFunctionsWithOriginalComment.xlsx");
    }
}