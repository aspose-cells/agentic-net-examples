// Title: Toggle Shape Texture Tiling in Excel with Aspose.Cells for .NET
// Description: A C# routine that loads or creates an Excel workbook, finds (or adds) a rectangle shape named "ToggleShape", applies a blue tissue‑paper texture fill, flips the Fill.TextureFill.IsTiling flag on each run, writes the new state to the console, and saves the file so the tiling setting persists.
// Keywords: Aspose.Cells texture tiling | C# shape texture fill | Excel shape IsTiling | toggle texture tiling Aspose.Cells | persist shape state Excel | macro‑like texture toggle | .NET Excel automation
// Common Searches: how to toggle texture tiling on a shape using Aspose.Cells C# | Aspose.Cells change IsTiling property of shape texture | persist shape texture tiling across workbook saves | add rectangle shape with texture fill Aspose.Cells | C# example for texture tiling toggle in Excel
// Developer Intent: Flip the IsTiling flag of a shape’s texture fill and keep the change in the workbook for subsequent executions.
// Use Cases: Automate a report that alternates between tiled and non‑tiled textures to emphasize sections. | Create a reusable macro‑like function that remembers a user‑selected tiling option. | Build an Excel‑based UI where a button toggles texture tiling without manual editing.
// AI Prompts: Generate a C# method that accepts a Worksheet and shape name, ensures the shape has a texture fill, and toggles its IsTiling property using Aspose.Cells. | Write Aspose.Cells code to add a rectangle with a custom texture, set an initial tiling state, invert the tiling flag on each run, and save the workbook. | Explain how to store and retrieve the texture tiling setting of a shape in an Excel file with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureToggle
{
    // A C# routine that loads or creates an Excel workbook, finds (or adds) a rectangle shape named "ToggleShape", applies a blue tissue‑paper texture fill, flips the Fill.TextureFill.IsTiling flag on each run, writes the new state to the console, and saves the file so the tiling setting persists.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that will store the shape and its tiling state
            string filePath = "TextureToggleDemo.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(filePath))
            {
                workbook = new Workbook(filePath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Work with the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Try to find a shape named "ToggleShape"
            Shape shape = null;
            foreach (Shape s in worksheet.Shapes)
            {
                if (s.Name == "ToggleShape")
                {
                    shape = s;
                    break;
                }
            }

            // If the shape does not exist, create it and set initial texture properties
            if (shape == null)
            {
                // Add a rectangle shape
                shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);
                shape.Name = "ToggleShape";

                // Configure texture fill
                shape.Fill.FillType = FillType.Texture;
                shape.Fill.TextureFill.Type = TextureType.BlueTissuePaper;

                // Start with tiling disabled
                shape.Fill.TextureFill.IsTiling = false;
            }

            // Toggle the IsTiling property
            bool currentTiling = shape.Fill.TextureFill.IsTiling;
            shape.Fill.TextureFill.IsTiling = !currentTiling;

            // Output the new state to the console
            Console.WriteLine($"Texture tiling is now set to: {shape.Fill.TextureFill.IsTiling}");

            // Save the workbook, preserving the toggled state for the next execution
            workbook.Save(filePath);
        }
    }
}
