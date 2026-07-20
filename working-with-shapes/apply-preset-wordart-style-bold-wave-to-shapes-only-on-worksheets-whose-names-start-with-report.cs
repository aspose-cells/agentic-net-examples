// Title: C# – Apply Bold Wave WordArt preset to WordArt shapes on worksheets whose names start with "Report" using Aspose.Cells
// Description: Loads an Excel workbook with Aspose.Cells, filters worksheets whose names begin with "Report", iterates through each shape, and for every WordArt object sets the TextEffect preset to Wave1 (Bold Wave) and makes the font bold. The modified workbook is then saved.
// Keywords: Aspose.Cells WordArt preset | C# apply Bold Wave WordArt | MsoPresetTextEffectShape.Wave1 | filter worksheets by name Aspose | Excel shape TextEffect formatting | batch update WordArt .NET | Excel automation WordArt style | apply bold font to WordArt | Aspose.Cells shape collection
// Common Searches: Aspose.Cells set WordArt preset to Wave1 in C# | apply Bold Wave WordArt only on worksheets starting with Report | filter Excel sheets by name and change WordArt style using Aspose | C# code to make all WordArt shapes bold in specific worksheets | batch modify WordArt text effect with Aspose.Cells
// Developer Intent: Apply the Bold Wave (Wave1) WordArt preset and bold font to every WordArt shape on worksheets whose names start with "Report".
// Use Cases: Standardize report titles by automatically applying a bold wave WordArt effect to all report sheets. | Enforce corporate branding across multiple workbooks by batch‑updating WordArt shapes on designated worksheets. | Prepare a template workbook where any new sheet prefixed with "Report" inherits the same WordArt styling without manual editing.
// AI Prompts: Generate C# code with Aspose.Cells that applies the Bold Wave WordArt preset and bold font to all WordArt shapes on worksheets whose names start with "Report" and saves the workbook. | Show how to modify the sample to also change the WordArt font color to blue while keeping the Bold Wave preset for those shapes. | Explain how to extend the example to target only WordArt shapes whose name contains "Title" in addition to the worksheet‑name filter.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook with Aspose.Cells, filters worksheets whose names begin with "Report", iterates through each shape, and for every WordArt object sets the TextEffect preset to Wave1 (Bold Wave) and makes the font bold. The modified workbook is then saved.
class ApplyBoldWaveWordArt
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Process only worksheets whose names start with "Report"
            if (worksheet.Name.StartsWith("Report", StringComparison.OrdinalIgnoreCase))
            {
                ShapeCollection shapes = worksheet.Shapes;

                // Loop through each shape in the worksheet
                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // Apply only to WordArt shapes
                    if (shape.IsWordArt)
                    {
                        // Access the TextEffectFormat of the WordArt shape
                        TextEffectFormat textEffect = shape.TextEffect;

                        // Set the preset shape to Wave1 (Bold Wave) and make the font bold
                        textEffect.PresetShape = MsoPresetTextEffectShape.Wave1;
                        textEffect.FontBold = true;
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
