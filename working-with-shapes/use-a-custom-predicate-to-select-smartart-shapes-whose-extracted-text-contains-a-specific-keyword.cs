using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtSelection
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Define the keyword to search for within SmartArt text
            string keyword = "TargetKeyword";

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the current worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Custom predicate: select only SmartArt shapes whose Text contains the keyword
                    if (shape.IsSmartArt && !string.IsNullOrEmpty(shape.Text) &&
                        shape.Text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Example action: output shape information
                        Console.WriteLine($"Found SmartArt shape on sheet '{sheet.Name}' with Id {shape.Id} and Name '{shape.Name}'.");
                        
                        // Additional processing can be performed here, e.g., modify the shape, extract data, etc.
                    }
                }
            }

            // Save the workbook (if any modifications were made)
            workbook.Save("output.xlsx");
        }
    }
}