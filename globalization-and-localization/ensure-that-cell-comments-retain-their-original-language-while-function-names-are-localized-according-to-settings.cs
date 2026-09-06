// Title: Keep cell comment language unchanged while localizing Excel formula function names with Aspose.Cells for .NET
// AI Prompts: Configure Workbook.Settings.CultureInfo to a target locale (e.g., fr-FR) and confirm that formulas show localized function names while comment text remains exactly as entered. | Insert a Japanese comment into a cell, apply French culture for formula localization, and save the workbook ensuring the comment content is not altered.
// Common Searches: Aspose.Cells preserve multilingual comments when setting CultureInfo for formula localization | C# change workbook culture to French to translate Excel functions without modifying cell comments | retain original language of cell comments after applying localized function names in Aspose.Cells
// Tags: localize Excel formulas Aspose.Cells CultureInfo | preserve Unicode cell comments Aspose.Cells | set workbook culture for function name translation .NET | multilingual comment retention in saved workbook | French function name localization Aspose.Cells

using Aspose.Cells;
using System.Globalization;

// The example loads a workbook, sets Workbook.Settings.CultureInfo to French so that formula functions appear in the French language, adds a Japanese comment to a cell that stays unchanged, and saves the file, demonstrating how to localize function names without affecting the original language of cell comments.
class Program
{
    static void Main()
    {
        // Load an existing workbook (load rule)
        var workbook = new Workbook("input.xlsx");

        // Localize function names by setting the desired culture (e.g., French)
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Example formula – when the workbook is opened, the function name will appear localized
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Formula = "=SUM(B1:B5)";

        // Add a comment in its original language (e.g., Japanese). The comment text is stored as‑is.
        var comment = sheet.Comments[sheet.Comments.Add("B2")];
        comment.Note = "これはテストコメントです。";

        // Save the workbook (save rule)
        workbook.Save("output.xlsx");
    }
}
