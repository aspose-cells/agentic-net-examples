// Title: Apply a consistent 5‑point shadow offset to all shapes in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Use the Aspose.Cells C# API to loop through every worksheet and each shape, enable the shape's Shadow, set OffsetX and OffsetY to 5 points, and then save the workbook. | Implement version‑safe shadow formatting by accessing Shape.Shadow dynamically and applying the same X/Y offset to every shape in an Excel file.
// Common Searches: C# Aspose.Cells set shadow offset for all shapes in a workbook | How to enable shape shadows and define X/Y offset in Excel with Aspose.Cells .NET | Iterate worksheets and shapes to apply uniform shadow using Aspose.Cells | Dynamic access to Shape.Shadow property for different Aspose.Cells versions | Apply same shadow offset to multiple Excel shapes with Aspose.Cells C#
// Tags: consistent shape shadow offset Aspose.Cells | shape shadow visibility .NET | iterate all worksheet shapes C# | version‑safe shadow property access Aspose.Cells | excel shape visual alignment

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads input.xlsx, iterates each worksheet and shape, makes the shadow visible, sets a 5‑point X and Y offset via dynamic property access for version safety, and saves the modified workbook to output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Uniform shadow offset values (in points)
            const double offsetX = 5.0;
            const double offsetY = 5.0;

            // Iterate through all worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Use dynamic to access Shadow property if it exists in the current Aspose.Cells version
                        dynamic dynShape = shape;
                        dynShape.Shadow.IsVisible = true;
                        dynShape.Shadow.OffsetX = offsetX;
                        dynShape.Shadow.OffsetY = offsetY;
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        // Shadow property not available in this version; safely ignore
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the application from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
