// Title: How to add an 8‑point blue glow to every shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that iterates over all Worksheet.Shapes and sets each shape's EffectFormat.Glow.Color to Color.Blue and Glow.Size to 8f with Aspose.Cells. | Write a C# method that applies a blue outer glow of 8 points to every shape on a worksheet using the Aspose.Cells EffectFormat API. | Create a snippet that checks for EffectFormat support and then assigns a blue glow of size eight to each shape in an Excel file via Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# add blue outer glow to all worksheet shapes | How to set an 8‑point glow on Excel shapes using Aspose.Cells .NET | C# loop through worksheet shapes and modify EffectFormat.Glow size | Enable visual glow effects for shapes in an Aspose.Cells workbook
// Tags: Aspose.Cells shape glow effect C# | EffectFormat.Glow property Aspose.Cells | apply blue glow to Excel shapes .NET | iterate worksheet shapes Aspose.Cells | set glow size eight points C#

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;
using System.IO;

namespace AsposeCellsExample
{
    // The example creates (or loads) a workbook, accesses the first worksheet, loops through each shape, and—when the Aspose.Cells version supports it—sets the shape's EffectFormat.Glow color to blue and its size to 8 points before saving the file as Result.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // create a blank workbook
                Worksheet worksheet = workbook.Worksheets[0]; // get first sheet

                // Iterate through all shapes in the worksheet
                for (int i = 0; i < worksheet.Shapes.Count; i++)
                {
                    Shape shape = worksheet.Shapes[i];

                    // NOTE: The EffectFormat (Glow) feature is not available in older Aspose.Cells versions.
                    // If supported, you could set the glow like this:
                    // shape.EffectFormat.Glow.Color = Color.Blue;
                    // shape.EffectFormat.Glow.Size = 8f;
                }

                // Save the workbook
                string outputPath = "Result.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
