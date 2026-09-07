// Title: Convert an Excel workbook to HTML while preserving every cell style using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as HTML with Aspose.Cells, ensuring all cell styles are kept by setting HtmlSaveOptions.ExcludeUnusedStyles to false. | Show how to export a workbook to HTML without removing unused styles using Aspose.Cells HtmlSaveOptions in a .NET application. | Create a C# program that converts Excel to HTML while disabling style pruning to compare full and reduced style sets.
// Common Searches: Aspose.Cells C# save workbook as HTML without excluding unused styles | How to keep all cell formatting when converting Excel to HTML with Aspose.Cells | HtmlSaveOptions.ExcludeUnusedStyles false example in .NET | Compare full style set vs reduced style set in Aspose.Cells HTML export
// Tags: Aspose.Cells HtmlSaveOptions ExcludeUnusedStyles | export Excel to HTML preserving styles | C# convert workbook to HTML with full style set | disable unused style removal Aspose.Cells | compare full versus reduced HTML style output

using Aspose.Cells;
using System;

// Loads input.xlsx, sets HtmlSaveOptions.ExcludeUnusedStyles to false to retain all cell styles, and saves the workbook as output_full_styles.html.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Set HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Disable exclusion of unused styles to retain the full style set
        htmlOptions.ExcludeUnusedStyles = false;

        // Save the workbook as HTML with the specified options
        workbook.Save("output_full_styles.html", htmlOptions);
    }
}
