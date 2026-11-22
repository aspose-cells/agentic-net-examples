// Title: Localize Excel formula names to French while keeping cell comment text unchanged using Aspose.Cells for .NET
// Common Searches: how to change Excel formula language to French with Aspose.Cells C# without modifying comments | preserve original comment language when setting workbook culture in Aspose.Cells | Aspose.Cells C# localize built‑in functions but keep comment notes intact | set CultureInfo for formula localization while retaining comment text in .NET | convert Excel formulas to French using Aspose.Cells while leaving comments unchanged
// Tags: Aspose.Cells formula localization using CultureInfo | retain original comment language after workbook culture change | C# localize built‑in Excel functions with Aspose.Cells | preserve cell comment notes when changing workbook culture | Excel workbook French function names Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The example loads an Excel workbook, sets workbook.Settings.CultureInfo to French to translate built‑in formula names, iterates through each worksheet's comments to demonstrate that comment text remains untouched, and saves the workbook, ensuring function names are localized while comment language is preserved.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Set the culture for function name localization (e.g., French)
        // This will localize built‑in function names while leaving comment text untouched.
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Iterate through all worksheets and their comments to demonstrate that
        // comment text remains unchanged (no modification is performed).
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Comment comment in sheet.Comments)
            {
                // Example: output comment text to console (optional)
                Console.WriteLine($"Comment in sheet '{sheet.Name}': {comment.Note}");
                // No changes are made to comment.Note, preserving the original language.
            }
        }

        // Save the workbook with localized function names (replace with desired output path)
        workbook.Save("Output.xlsx");
    }
}
