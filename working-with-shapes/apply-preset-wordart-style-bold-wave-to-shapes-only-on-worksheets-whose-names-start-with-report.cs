// Title: Aspose.Cells C# – Apply Bold Wave WordArt to Worksheets Prefixed with "Report"
// Description: Creates a workbook, adds worksheets, inserts WordArt shapes, then filters sheets whose names start with "Report" and sets each WordArt's TextEffectFormat to the Wave1 preset with bold font before saving the file.
// Keywords: Aspose.Cells WordArt style | C# WordArt Wave1 | bold WordArt preset | filter worksheets by name | TextEffectFormat C# | apply WordArt to specific sheets
// Common Searches: Aspose.Cells set WordArt bold wave on selected sheets | C# filter worksheets by prefix and change WordArt style | apply preset WordArt shape to Excel sheets using Aspose.Cells | how to make WordArt bold and wavy in .NET
// Developer Intent: Apply a bold wave WordArt effect to every WordArt shape on worksheets whose names begin with "Report".
// Use Cases: Standardize report sheet headings with a bold wave WordArt for brand consistency. | Automatically style newly generated monthly report tabs without manual editing. | Update legacy workbooks so only report‑named sheets receive the enhanced WordArt formatting.
// AI Prompts: Write C# code with Aspose.Cells that applies the Wave1 WordArt preset and bold font to all WordArt shapes on worksheets whose names start with "Report". | Show an Aspose.Cells example that filters worksheets by a name prefix and modifies TextEffectFormat for WordArt objects. | Explain how to change the preset shape to a different WordArt style while still targeting only "Report" worksheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtStyleExample
{
    // Creates a workbook, adds worksheets, inserts WordArt shapes, then filters sheets whose names start with "Report" and sets each WordArt's TextEffectFormat to the Wave1 preset with bold font before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add worksheets with names to demonstrate the filter
            Worksheet ws1 = workbook.Worksheets[workbook.Worksheets.Add()];
            ws1.Name = "Report_January";
            Worksheet ws2 = workbook.Worksheets[workbook.Worksheets.Add()];
            ws2.Name = "Data_Sheet";

            // Add a WordArt shape to each worksheet for testing
            ws1.Shapes.AddWordArt(PresetWordArtStyle.WordArtStyle1, "Sample", 2, 0, 2, 0, 100, 300);
            ws2.Shapes.AddWordArt(PresetWordArtStyle.WordArtStyle1, "Sample", 2, 0, 2, 0, 100, 300);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Process only worksheets whose names start with "Report"
                if (sheet.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
                {
                    // Iterate through all shapes in the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Apply only to WordArt shapes
                        if (shape.IsWordArt)
                        {
                            // Access the TextEffectFormat of the shape
                            TextEffectFormat textEffect = shape.TextEffect;

                            // Set the preset shape to Wave1 (represents a wave effect)
                            textEffect.PresetShape = MsoPresetTextEffectShape.Wave1;

                            // Make the text bold
                            textEffect.FontBold = true;
                        }
                    }
                }
            }

            // Save the workbook
            workbook.Save("ReportWordArtBoldWave.xlsx");
        }
    }
}
