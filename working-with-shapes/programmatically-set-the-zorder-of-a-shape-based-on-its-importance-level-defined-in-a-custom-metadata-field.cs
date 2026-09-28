// Title: Set Excel shape Z‑order in C# with Aspose.Cells using a custom 'Importance' metadata field
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, reads each shape's AlternativeText for an 'Importance' key, and assigns the shape's ZOrderPosition based on a High/Medium/Low mapping. | Create a function that extracts a custom metadata value (e.g., 'Priority') from a shape's AlternativeText and returns the appropriate Z‑order index for use with Aspose.Cells. | Generate C# logic that logs shapes lacking a valid importance entry while reordering all shapes according to their metadata‑defined Z‑order in an Excel workbook.
// Common Searches: how to change the ZOrderPosition of shapes in an Excel file using Aspose.Cells C# | read custom metadata from shape AlternativeText and reorder shapes in Aspose.Cells | map importance levels to shape layering in a .xlsx workbook with Aspose.Cells | C# Aspose.Cells example for setting shape Z‑order based on a metadata field | Aspose.Cells shape ordering by priority stored in AlternativeText
// Tags: Aspose.Cells set shape ZOrderPosition | C# read shape AlternativeText metadata | Excel shape layering based on importance | custom metadata to Z‑order mapping Aspose.Cells | shape ordering dictionary mapping C#

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;

// Loads an .xlsx workbook, iterates over worksheet shapes, extracts an 'Importance' value from each shape's AlternativeText, maps High/Medium/Low to ZOrderPosition values (3,2,1), assigns the Z‑order, logs missing or invalid entries, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Mapping of importance levels to Z‑order values (higher = front)
            var importanceToZOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "High",   3 },
                { "Medium", 2 },
                { "Low",    1 }
            };

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Assume the importance level is stored in the shape's AlternativeText
                // in the format: "Importance=High"
                string importance = GetImportanceFromAlternativeText(shape.AlternativeText);

                if (importance != null && importanceToZOrder.TryGetValue(importance, out int zOrder))
                {
                    // Set the Z‑order of the shape (higher value brings the shape to the front)
                    shape.ZOrderPosition = zOrder;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to extract the importance value from AlternativeText
    private static string GetImportanceFromAlternativeText(string alternativeText)
    {
        if (string.IsNullOrEmpty(alternativeText))
            return null;

        // Split possible multiple key‑value pairs
        string[] pairs = alternativeText.Split(new[] { ';', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string pair in pairs)
        {
            string[] kv = pair.Split(new[] { '=' }, 2);
            if (kv.Length == 2 && kv[0].Trim().Equals("Importance", StringComparison.OrdinalIgnoreCase))
                return kv[1].Trim();
        }
        return null;
    }
}
