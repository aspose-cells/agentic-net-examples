// Title: Keep cell comment language unchanged while localizing Excel function names with Aspose.Cells for .NET
// Description: Shows how to add a Japanese comment to a cell, configure SettableGlobalizationSettings to map the English function SUM to the French name SOMME (bidirectional), use the localized formula, calculate the workbook, and confirm that the comment’s author and text stay in Japanese before saving the file.
// Keywords: Aspose.Cells | .NET | cell comments localization | multilingual comments | SettableGlobalizationSettings | function name mapping | SUM to SOMME | Excel formula localization | Japanese comment | French function name | preserve comment language
// Common Searches: Aspose.Cells keep comment language after globalization | map Excel function SUM to French SOMME in C# | preserve Japanese comment in Aspose.Cells workbook | bidirectional function name localization Aspose.Cells | C# example SettableGlobalizationSettings localized formulas
// Developer Intent: The developer needs to retain the original language of cell comments while applying global settings that translate Excel function names to localized equivalents.
// Use Cases: Add comments in any language (e.g., Japanese) and ensure they remain unchanged after enabling globalization features. | Translate standard Excel functions such as SUM to a target language (e.g., French SOMME) and use them in formulas that calculate correctly. | Create workbooks that combine multilingual annotations with localized formulas for international users. | Generate XLSX files that respect both comment language integrity and localized function naming.
// AI Prompts: Write a C# Aspose.Cells example that inserts a Japanese comment, sets SettableGlobalizationSettings to map SUM to SOMME (bidirectional), applies the French formula, calculates the workbook, and verifies the comment language. | Provide code to map multiple Excel functions to their localized names using SettableGlobalizationSettings while keeping all cell comments in their original scripts. | Explain how to test that cell comments retain their original language after applying global function name localization in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCommentAndLocalizationDemo
{
    // Shows how to add a Japanese comment to a cell, configure SettableGlobalizationSettings to map the English function SUM to the French name SOMME (bidirectional), use the localized formula, calculate the workbook, and confirm that the comment’s author and text stay in Japanese before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 1. Add a comment in a specific language (e.g., Japanese)
            // ------------------------------------------------------------
            int commentIndex = sheet.Comments.Add("C3");               // Add comment to cell C3
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "これはテストコメントです。";               // Japanese comment text
            comment.Author = "山田太郎";                               // Japanese author name
            comment.IsVisible = true;                                 // Make comment visible

            // ------------------------------------------------------------
            // 2. Configure globalization settings to localize function names
            // ------------------------------------------------------------
            // Create SettableGlobalizationSettings instance
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Map the standard English function name "SUM" to a French localized name "SOMME"
            // The third parameter 'true' makes the mapping bidirectional
            globalization.SetLocalFunctionName("SUM", "SOMME", true);

            // Apply the settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // ------------------------------------------------------------
            // 3. Use the localized function name in a formula
            // ------------------------------------------------------------
            // Populate some data for the formula
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Use the localized function name "SOMME" in the formula
            sheet.Cells["B1"].Formula = "=SOMME(A1:A3)";

            // Calculate formulas so that the result is stored in the cell
            workbook.CalculateFormula();

            // ------------------------------------------------------------
            // 4. Verify that the comment retains its original language
            // ------------------------------------------------------------
            Console.WriteLine("Comment Author (should be Japanese): " + comment.Author);
            Console.WriteLine("Comment Text (should be Japanese): " + comment.Note);
            Console.WriteLine("Formula Result (using localized function): " + sheet.Cells["B1"].Value);

            // ------------------------------------------------------------
            // 5. Save the workbook
            // ------------------------------------------------------------
            workbook.Save("CommentAndLocalizedFunctionDemo.xlsx");
        }
    }
}
