// Title: Apply Uniform Orange Glow to Shapes Named "Important" Using Aspose.Cells C#
// Description: C# script that loads an Excel workbook (or creates one), scans every worksheet, and adds a 12‑point, 40% transparent orange glow to each shape whose Name contains the word "Important". The modified file is saved as Output_With_Glow.xlsx.
// Keywords: Aspose.Cells C# glow effect | Excel shape glow Aspose | filter shapes by name Aspose.Cells | apply orange glow to shapes | shape formatting Excel .NET | uniform glow effect workbook
// Common Searches: how to add glow to specific shapes in Excel using Aspose.Cells | C# apply orange glow to shapes containing 'Important' | iterate worksheet shapes and set glow properties Aspose.Cells | Aspose.Cells set shape glow color and transparency
// Developer Intent: Programmatically add the same orange glow to every shape whose Name includes "Important" in an Excel workbook.
// Use Cases: Highlight critical diagram elements in financial reports. | Emphasize warning icons on dashboard worksheets. | Standardize visual cues across multiple sheets by auto‑applying a glow to important shapes.
// AI Prompts: Generate C# code with Aspose.Cells to apply a red 10‑point glow to shapes whose name contains "Alert". | Rewrite the glow‑application loop using LINQ to filter shapes by a keyword. | Explain how to set glow color and transparency dynamically based on each shape's type.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExamples
{
    // C# script that loads an Excel workbook (or creates one), scans every worksheet, and adds a 12‑point, 40% transparent orange glow to each shape whose Name contains the word "Important". The modified file is saved as Output_With_Glow.xlsx.
    public class ApplyUniformGlowToImportantShapes
    {
        public static void Run()
        {
            try
            {
                Workbook workbook;
                string inputPath = "input.xlsx";

                // Load existing workbook if it exists; otherwise create a new workbook
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    ShapeCollection shapes = sheet.Shapes;

                    // Process each shape
                    for (int i = 0; i < shapes.Count; i++)
                    {
                        Shape shape = shapes[i];

                        // Apply glow to shapes whose name contains "Important"
                        if (!string.IsNullOrEmpty(shape.Name) && shape.Name.Contains("Important"))
                        {
                            GlowEffect glow = shape.Glow;
                            glow.Size = 12;               // radius in points
                            glow.Transparency = 0.4;      // 40% transparent

                            // Set a uniform orange glow color
                            CellsColor glowColor = workbook.CreateCellsColor();
                            glowColor.Color = Color.Orange;
                            glow.Color = glowColor;
                        }
                    }
                }

                // Save the workbook with applied effects
                string outputPath = "Output_With_Glow.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyUniformGlowToImportantShapes.Run();
        }
    }
}
