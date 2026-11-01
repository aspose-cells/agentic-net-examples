// Title: Localize Excel formula function names to French while keeping cell comments unchanged using Aspose.Cells for .NET
// Common Searches: how to change Excel formula language to French with Aspose.Cells C# without affecting existing comments | preserve original comment text when setting workbook culture in Aspose.Cells | Aspose.Cells C# translate function names based on CultureInfo while retaining comment language | keep cell comments intact while localizing formulas in an Excel file using Aspose.Cells | set workbook Settings.CultureInfo to fr-FR and avoid comment translation in Aspose.Cells
// Tags: Workbook.Settings.CultureInfo formula localization | retain cell comment content during culture change | French function name translation Aspose.Cells | comment integrity with Excel workbook localization | C# Aspose.Cells preserve comments while localizing formulas

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, sets Workbook.Settings.CultureInfo to French to automatically translate formula function names, iterates through every cell to reassign each comment's Note property to its original text ensuring comments remain unchanged, and saves the workbook with localized functions while preserving the original comment language.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Set the culture for formula translation (e.g., French).
            // This localizes function names according to the specified culture.
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Preserve original comment text (no translation or modification).
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Cell cell in sheet.Cells)
                {
                    // In Aspose.Cells, a comment is accessed via the Comment property.
                    if (cell.Comment != null)
                    {
                        // Retrieve the existing comment text.
                        string originalComment = cell.Comment.Note;
                        // Reassign to ensure it remains unchanged.
                        cell.Comment.Note = originalComment;
                    }
                }
            }

            // Save the workbook with localized function names while keeping comments unchanged.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
