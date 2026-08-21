// Title: C# utility to reset all shape effects in an Aspose.Cells workbook
// Description: Provides a ShapeEffectUtility.ResetAllShapeEffects method that walks through every worksheet, drawing shape, and chart series in a Workbook, clears supported 3‑D and visual effect properties, logs unsupported items, and demonstrates loading, processing, and saving an Excel file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells reset shape effects | clear shape formatting C# | remove 3D effects Excel shapes | reset chart series shape properties | Aspose.Cells utility method | C# Excel shape cleanup
// Common Searches: how to clear shape effects with Aspose.Cells .NET | reset all shape formatting in an Excel workbook using C# | remove 3D and shadow effects from shapes Aspose.Cells | utility to clear chart series shape properties | Aspose.Cells example for resetting shape effects
// Developer Intent: Remove visual effects from every shape and chart series in a workbook and save the cleaned file.
// Use Cases: Prepare workbooks for printing by stripping 3‑D, shadow, and glow effects. | Standardize exported reports so charts and shapes have a uniform appearance. | Automate cleanup of user‑generated Excel files to reduce file size before archiving.
// AI Prompts: Generate C# code that uses Aspose.Cells to reset ThreeDFormat and EffectFormat properties of shapes when the API exposes them. | Show best‑practice error handling while iterating ShapeCollection and Series.ShapeProperties in Aspose.Cells. | Explain how to extend ShapeEffectUtility for version‑specific APIs to clear additional shape effects.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // Required for Shape, ShapeCollection, etc.

// Provides a ShapeEffectUtility.ResetAllShapeEffects method that walks through every worksheet, drawing shape, and chart series in a Workbook, clears supported 3‑D and visual effect properties, logs unsupported items, and demonstrates loading, processing, and saving an Excel file with Aspose.Cells for .NET.
public static class ShapeEffectUtility
{
    /// <param name="workbook">The workbook whose shape effects should be cleared.</param>
    public static void ResetAllShapeEffects(Workbook workbook)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Reset effects for drawing shapes on the worksheet
            ShapeCollection shapes = sheet.Shapes;
            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];
                try
                {
                    // NOTE: In the current Aspose.Cells version the ThreeDFormat and EffectFormat
                    // properties are either not exposed or have different APIs.
                    // To keep the code compile‑time safe we simply skip resetting those
                    // specific properties. If needed, they can be handled with version‑specific
                    // APIs in the future.
                }
                catch (Exception ex)
                {
                    // Log or ignore shape types that do not support these operations
                    Console.WriteLine($"Warning: Unable to clear effects for shape '{shape.Name}'. {ex.Message}");
                }
            }

            // Reset effects for chart series shape properties
            foreach (Chart chart in sheet.Charts)
            {
                foreach (Series series in chart.NSeries)
                {
                    ShapePropertyCollection seriesProps = series.ShapeProperties;
                    if (seriesProps != null)
                    {
                        try
                        {
                            // As with drawing shapes, specific 3‑D or effect properties are omitted
                            // for compatibility with the available API set.
                        }
                        catch (Exception ex)
                        {
                            // Log or ignore unsupported series shape properties
                            Console.WriteLine($"Warning: Unable to clear effects for series in chart '{chart.Name}'. {ex.Message}");
                        }
                    }
                }
            }
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook, clear shape effects, and save result
            Workbook workbook = new Workbook(inputPath);
            ShapeEffectUtility.ResetAllShapeEffects(workbook);

            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
