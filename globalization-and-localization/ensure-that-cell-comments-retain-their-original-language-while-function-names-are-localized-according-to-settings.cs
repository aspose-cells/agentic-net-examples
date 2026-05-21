using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom globalization settings that only localize function names
    public class CustomGlobalizationSettings : SettableGlobalizationSettings
    {
        // No additional overrides needed for comment titles
    }

    public class CommentAndFunctionLocalizationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare sample data for the formula
                worksheet.Cells["A1"].PutValue(10);
                worksheet.Cells["A2"].PutValue(20);
                worksheet.Cells["A3"].PutValue(30);

                // Create custom globalization settings
                CustomGlobalizationSettings settings = new CustomGlobalizationSettings();

                // Map the standard SUM function to its Spanish equivalent "SUMA"
                settings.SetLocalFunctionName("SUM", "SUMA", true);

                // Apply the settings to the workbook
                workbook.Settings.GlobalizationSettings = settings;

                // Add a comment in cell B2 with Spanish content
                int commentIndex = worksheet.Comments.Add("B2");
                Comment comment = worksheet.Comments[commentIndex];
                comment.Note = "Este es un comentario en español.";
                comment.IsVisible = true;

                // Set the comment author (title) directly
                comment.Author = "Comentario de Celda";

                // Use the localized function name in a formula
                worksheet.Cells["B2"].Formula = "=SUMA(A1:A3)";

                // Calculate formulas
                workbook.CalculateFormula();

                // Output results to console
                Console.WriteLine($"Formula result in B2: {worksheet.Cells["B2"].Value}");
                Console.WriteLine($"Comment author (title): {comment.Author}");
                Console.WriteLine($"Comment note: {comment.Note}");

                // Save the workbook
                string outputPath = "CommentAndFunctionLocalizationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            CommentAndFunctionLocalizationDemo.Run();
        }
    }
}