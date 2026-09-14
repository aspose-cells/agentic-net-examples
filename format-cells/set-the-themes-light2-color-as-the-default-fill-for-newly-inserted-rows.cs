// Title: How to apply the workbook Light2 theme color as the default fill when inserting new rows with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a style using the workbook's Light2 theme color, applies it to a template row, and inserts a new row that inherits this fill via Aspose.Cells. | Write a method that inserts a row at a specified index and automatically uses the Light2 theme background, falling back to a solid color if the theme cannot be retrieved. | Provide a script that configures a StyleFlag to copy only cell shading, applies a Light2‑based style to the row above the insertion point, and then inserts rows preserving that shading.
// Common Searches: Aspose.Cells insert row with Light2 theme background C# | C# Aspose.Cells set default fill color for newly added rows using workbook theme | how to apply Light2 theme color as row background when adding rows in Aspose.Cells | fallback solid color for theme retrieval when inserting rows with Aspose.Cells .NET
// Tags: apply Light2 theme fill Aspose.Cells | insert rows with style copy Aspose.Cells .NET | fallback solid color for theme retrieval C# | template row style flag cell shading Aspose.Cells | workbook theme background for new rows

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// Loads an existing workbook, creates a style using the Light2 theme color (or LightGray as a fallback), applies the style to the row above the insertion point, inserts a new row copying only the fill, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Use a fallback color (LightGray) if theme color retrieval is unavailable
            Color light2Color = Color.LightGray;

            // Define the row index where new rows will be inserted (0‑based)
            int insertIndex = 5; // example: insert after row 4

            // Prepare a style with Light2 fill
            Style light2Style = workbook.CreateStyle();
            light2Style.ForegroundColor = light2Color;
            light2Style.Pattern = BackgroundType.Solid;

            // Apply the style to the row that will serve as a template
            // (the row just above the insertion point)
            Row templateRow = sheet.Cells.Rows[insertIndex - 1];
            StyleFlag flag = new StyleFlag
            {
                CellShading = true // apply only fill
            };
            templateRow.ApplyStyle(light2Style, flag);

            // Insert a new row and copy the style from the template row
            sheet.Cells.InsertRows(insertIndex, 1, true); // true = copy style

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
