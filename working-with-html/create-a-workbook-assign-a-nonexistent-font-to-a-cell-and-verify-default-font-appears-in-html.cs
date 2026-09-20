// Title: Generate an Excel workbook in C#, assign a non‑existent font to a cell, and confirm default font substitution in the HTML export using Aspose.Cells
// AI Prompts: Write C# code that creates a new Workbook with Aspose.Cells, sets the Font.Name of cell A1 to a font that does not exist on the system, saves the workbook as HTML, reads the HTML file, and asserts that the missing font name is not present. | Provide a .NET snippet demonstrating how Aspose.Cells replaces an unavailable font with the default font during HTML conversion, including verification logic that checks the generated HTML content.
// Common Searches: how does Aspose.Cells handle missing fonts when exporting Excel to HTML in C# | C# Aspose.Cells verify default font substitution in HTML output | example of setting a non‑existent font on a cell and checking HTML result with Aspose.Cells | Aspose.Cells HTML export fallback to Arial for unavailable fonts
// Tags: Aspose.Cells set cell font C# | HTML export font fallback Aspose.Cells | verify default font in generated HTML | nonexistent font handling Aspose.Cells | Workbook.Save Html Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Creates a workbook, applies a non‑existent font to cell A1, saves as HTML, reads the HTML file, and checks that the missing font name is absent, confirming Aspose.Cells substitutes the default font.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and the first cell (A1)
        Worksheet sheet = workbook.Worksheets[0];
        Cell cell = sheet.Cells["A1"];

        // Assign a non‑existent font to the cell
        Style style = cell.GetStyle();
        style.Font.Name = "NonExistentFontXYZ"; // This font does not exist on the system
        cell.SetStyle(style);
        cell.PutValue("Test text");

        // Save the workbook as HTML
        string htmlPath = "output.html";
        workbook.Save(htmlPath, SaveFormat.Html);

        // Load the generated HTML
        string htmlContent = File.ReadAllText(htmlPath);

        // Verify that the non‑existent font name does NOT appear in the HTML.
        // Aspose.Cells substitutes the missing font with the default font (usually Arial).
        bool containsMissingFont = htmlContent.IndexOf("NonExistentFontXYZ", StringComparison.OrdinalIgnoreCase) >= 0;

        if (containsMissingFont)
        {
            Console.WriteLine("The non‑existent font name was found in the HTML. Verification failed.");
        }
        else
        {
            Console.WriteLine("The non‑existent font name was NOT found in the HTML. Default font is applied. Verification succeeded.");
        }
    }
}
