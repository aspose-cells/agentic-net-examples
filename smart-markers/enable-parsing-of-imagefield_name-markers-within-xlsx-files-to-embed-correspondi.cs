using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImageMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the template workbook that contains {{Image:field_name}} markers
            string templatePath = "template.xlsx";

            // Load the workbook (lifecycle: load)
            Workbook workbook = new Workbook(templatePath);

            // Mapping of field names to image file paths
            // In a real scenario this could come from a database, API, etc.
            var imageMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Logo", "logo.png" },
                { "Photo", "photo.jpg" },
                { "Signature", "signature.png" }
            };

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all used cells
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Only consider string cells that may contain the marker
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue?.Trim();

                            // Check for the marker pattern {{Image:field_name}}
                            if (!string.IsNullOrEmpty(text) &&
                                text.StartsWith("{{Image:", StringComparison.Ordinal) &&
                                text.EndsWith("}}", StringComparison.Ordinal))
                            {
                                // Extract the field name between the braces
                                int startIdx = "{{Image:".Length;
                                int length = text.Length - startIdx - 2; // subtract closing "}}"
                                string fieldName = text.Substring(startIdx, length).Trim();

                                // Look up the image file path
                                if (imageMap.TryGetValue(fieldName, out string imagePath) && File.Exists(imagePath))
                                {
                                    // Read image bytes
                                    byte[] imageBytes = File.ReadAllBytes(imagePath);

                                    // Embed the image into the cell
                                    cell.EmbeddedImage = imageBytes;

                                    // Optionally clear the placeholder text
                                    cell.PutValue(string.Empty);
                                }
                                else
                                {
                                    Console.WriteLine($"Image for field '{fieldName}' not found.");
                                }
                            }
                        }
                    }
                }
            }

            // Save the modified workbook (lifecycle: save)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}