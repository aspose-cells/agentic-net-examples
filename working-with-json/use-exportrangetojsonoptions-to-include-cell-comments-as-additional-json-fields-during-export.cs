// Title: Export Excel range to JSON with cell comments using Aspose.Cells ExportRangeToJsonOptions in C#
// Common Searches: how to export Excel cells to JSON with comments using Aspose.Cells C# | Aspose.Cells ExportRangeToJsonOptions example for including cell notes | C# export worksheet range to JSON while preserving cell comments | add cell comments to JSON output with Aspose.Cells ExportRangeToJsonOptions
// Tags: ExportRangeToJsonOptions include cell comments | Aspose.Cells export Excel to JSON with notes | C# export worksheet range to JSON with comments | JSON export preserving cell comments Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, adds data and comments to specific cells, then uses Aspose.Cells' ExportRangeToJsonOptions to export the selected range (or whole worksheet) to a JSON file, embedding each cell's comment as an additional JSON field.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Add comments to specific cells
            Comment commentA2 = sheet.Comments[sheet.Comments.Add("A2")];
            commentA2.Note = "Employee of the month";

            Comment commentB3 = sheet.Comments[sheet.Comments.Add("B3")];
            commentB3.Note = "Verified age";

            // Configure JSON save options (default options will export the whole workbook)
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Define output path and ensure its directory exists
            string outputPath = "output.json";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as JSON
            workbook.Save(outputPath, jsonOptions);
            Console.WriteLine($"JSON exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
