// Title: Export Excel worksheet cells to JSON with values, data types and style details using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, iterates over all used cells in a worksheet, captures each cell's raw value, its data type, and selected style properties (font name, size, bold, italic, font color, background color, horizontal and vertical alignment), and serializes the result to a formatted JSON file. | Extend the export to also include each cell's formula text and a flag indicating whether the cell belongs to a merged range, while preserving the existing JSON structure. | Replace the color conversion routine so that exported colors are CSS‑compatible RGB strings (e.g., "rgb(17,34,51)") instead of ARGB hex values.
// Common Searches: how to convert Excel worksheet to JSON with cell formatting using Aspose.Cells C# | Aspose.Cells export cell style information to JSON file | C# serialize Excel cell data type and font attributes to JSON | include merged cells and formulas in Aspose.Cells JSON export | convert System.Drawing.Color to CSS rgb string in Aspose.Cells JSON output
// Tags: Aspose.Cells JSON export of cell styles | C# serialize Excel cell values and types | extract font and background colors Aspose.Cells | include cell alignment data in JSON | handle merged cells Aspose.Cells export

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsJsonExport
{
    // Represents the serialized information of a single cell.
    // The program loads an Excel workbook with Aspose.Cells, iterates through the used range of the first worksheet, captures each cell's value, data type, and a subset of style attributes (font name, size, bold, italic, font and background colors, horizontal and vertical alignment), converts colors to ARGB hex, and writes the collected rows as an indented JSON array to a specified file.
    public class CellInfo
    {
        // The raw value of the cell.
        public object Value { get; set; }

        // The data type of the cell (e.g., String, Double, DateTime, Boolean, etc.).
        public string DataType { get; set; }

        // A simplified representation of the cell's style.
        public CellStyleInfo Style { get; set; }
    }

    // Represents a subset of style properties we want to include in the JSON.
    public class CellStyleInfo
    {
        public string FontName { get; set; }
        public double FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public string FontColor { get; set; }          // ARGB hex string
        public string BackgroundColor { get; set; }    // ARGB hex string
        public string HorizontalAlignment { get; set; }
        public string VerticalAlignment { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source Excel file.
                string excelPath = @"C:\Data\Sample.xlsx";

                // Path where the resulting JSON will be saved.
                string jsonPath = @"C:\Data\SampleExport.json";

                // Verify that the source file exists.
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                    return;
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(jsonPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook.
                Workbook workbook = new Workbook(excelPath);

                // Get the first worksheet (you can change the index or name as needed).
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the used range.
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Prepare a list to hold rows; each row is a list of CellInfo objects.
                var rows = new List<List<CellInfo>>();

                // Iterate through each row.
                for (int row = 0; row <= maxRow; row++)
                {
                    var currentRow = new List<CellInfo>();

                    // Iterate through each column in the current row.
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Build the style information.
                        Style style = cell.GetStyle();
                        var styleInfo = new CellStyleInfo
                        {
                            FontName = style.Font.Name,
                            FontSize = style.Font.Size,
                            Bold = style.Font.IsBold,
                            Italic = style.Font.IsItalic,
                            FontColor = ColorToHex(style.Font.Color),
                            BackgroundColor = ColorToHex(style.BackgroundColor),
                            HorizontalAlignment = style.HorizontalAlignment.ToString(),
                            VerticalAlignment = style.VerticalAlignment.ToString()
                        };

                        // Build the cell information.
                        var cellInfo = new CellInfo
                        {
                            Value = cell.Value,
                            DataType = cell.Type.ToString(),
                            Style = styleInfo
                        };

                        currentRow.Add(cellInfo);
                    }

                    rows.Add(currentRow);
                }

                // Serialize the rows collection to JSON with indentation for readability.
                string json = JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented = true });

                // Write the JSON string to the output file.
                File.WriteAllText(jsonPath, json);

                Console.WriteLine($"Worksheet data exported to JSON at: {jsonPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to convert a System.Drawing.Color to an ARGB hex string (e.g., "#FF112233").
        private static string ColorToHex(Color color)
        {
            // If the color is empty (no color), return null.
            if (color.IsEmpty)
                return null;

            return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
        }
    }
}
