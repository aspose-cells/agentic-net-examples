using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook that was saved as HTML.
        // The constructor automatically detects the format from the file extension.
        Workbook workbook = new Workbook("input.html");

        // Configure JSON export options to preserve formatting metadata and structure.
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            // Keep style information (fonts, colors, borders, etc.) for each cell.
            ExportStylePool = true,

            // Export the workbook as a nested hierarchical JSON structure.
            ExportNestedStructure = true,

            // Include empty cells in the output (as null) to retain the exact layout.
            ExportEmptyCells = true,

            // Export cell values as strings to preserve the displayed text.
            ExportAsString = true
        };

        // Save the workbook as a JSON file using the configured options.
        workbook.Save("output.json", jsonOptions);
    }
}