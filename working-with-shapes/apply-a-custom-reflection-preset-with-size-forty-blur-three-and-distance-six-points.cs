using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load a template if it exists)
            Workbook workbook;
            string templatePath = "Template.xlsx";
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, row offset, column offset, width, height
            Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 100, 100);

            // Configure reflection effect (the property is read‑only; modify its members directly)
            shape.Reflection.Type = ReflectionEffectType.Custom; // custom preset
            shape.Reflection.Size = 40;      // size in percentage
            shape.Reflection.Blur = 3;       // blur radius in points
            shape.Reflection.Distance = 6;   // distance in points

            // Save the workbook with the applied reflection effect
            string outputPath = "CustomReflection.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}