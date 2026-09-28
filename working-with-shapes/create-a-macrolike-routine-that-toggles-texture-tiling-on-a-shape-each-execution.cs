// Title: How to toggle texture tiling (Tile vs Stretch) of a named shape in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate a C# method that takes a workbook file path and a shape name, finds the shape, and switches its FillFormat texture alignment between Tile and Stretch with Aspose.Cells. | Write C# code that uses reflection to detect and toggle either the TextureAlignment property or the TextureScaleX/TextureScaleY properties of a shape's FillFormat for version‑agnostic texture tiling control. | Create robust error‑handling for a routine that validates the file, locates the shape by name, toggles its texture tiling, and saves the workbook back to the same .xlsx file.
// Common Searches: aspnet toggle shape texture tiling in excel using aspose.cells | c# change shape fill texture alignment tile stretch aspose cells | how to use reflection to modify FillFormat texture properties in Aspose.Cells | toggle texture scaling of a shape programmatically in an .xlsx file | asp.net find shape by name and change its texture mode with Aspose.Cells
// Tags: aspocells shape texture tiling toggle | c# fillformat texturealignment reflection | excel shape fill tile stretch | aspocells version‑agnostic texture scaling | c# toggle shape fill texture in xlsx

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Reflection;

namespace ShapeTextureTilerApp
{
    // Provides a C# routine that locates a named shape in the first worksheet of an Excel file, toggles its texture tiling between Tile and Stretch using reflection for cross‑version compatibility, and saves the workbook.
    public class ShapeTextureTiler
    {
        // Toggles the texture tiling property of a shape in the first worksheet.
        public static void ToggleTextureTiling(string filePath, string shapeName)
        {
            try
            {
                if (!File.Exists(filePath))
                    throw new FileNotFoundException($"Workbook file not found: {filePath}");

                // Load the workbook.
                Workbook workbook = new Workbook(filePath);
                Worksheet sheet = workbook.Worksheets[0];

                // Locate the shape by name.
                Shape targetShape = null;
                foreach (Shape shape in sheet.Shapes)
                {
                    if (!string.IsNullOrEmpty(shape.Name) &&
                        shape.Name.Equals(shapeName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetShape = shape;
                        break;
                    }
                }

                if (targetShape == null)
                    throw new ArgumentException($"Shape '{shapeName}' not found in the worksheet.");

                // Attempt to toggle texture tiling.
                FillFormat fill = targetShape.Fill;

                // Use reflection to access TextureAlignment if it exists in the current Aspose.Cells version.
                PropertyInfo texAlignProp = typeof(FillFormat).GetProperty("TextureAlignment", BindingFlags.Public | BindingFlags.Instance);
                if (texAlignProp != null && texAlignProp.CanRead && texAlignProp.CanWrite)
                {
                    object currentValue = texAlignProp.GetValue(fill);
                    Type enumType = texAlignProp.PropertyType;

                    // Determine the enum names for Tile and Stretch (case‑insensitive).
                    object tileValue = Enum.Parse(enumType, "Tile", true);
                    object stretchValue = Enum.Parse(enumType, "Stretch", true);

                    // Toggle between Tile and Stretch.
                    object newValue = currentValue.Equals(tileValue) ? stretchValue : tileValue;
                    texAlignProp.SetValue(fill, newValue);
                }
                else
                {
                    // Fallback: try to toggle TextureScaleX/Y via reflection if they exist.
                    PropertyInfo scaleXProp = typeof(FillFormat).GetProperty("TextureScaleX", BindingFlags.Public | BindingFlags.Instance);
                    PropertyInfo scaleYProp = typeof(FillFormat).GetProperty("TextureScaleY", BindingFlags.Public | BindingFlags.Instance);

                    if (scaleXProp != null && scaleYProp != null && scaleXProp.CanRead && scaleXProp.CanWrite && scaleYProp.CanRead && scaleYProp.CanWrite)
                    {
                        double scaleX = Convert.ToDouble(scaleXProp.GetValue(fill));
                        double scaleY = Convert.ToDouble(scaleYProp.GetValue(fill));

                        // When both scales are 0 the texture is stretched; any positive value repeats the texture.
                        bool isStretched = scaleX == 0 && scaleY == 0;
                        double newScale = isStretched ? 1 : 0;

                        scaleXProp.SetValue(fill, newScale);
                        scaleYProp.SetValue(fill, newScale);
                    }
                    // If neither property exists, no further action is required.
                }

                // Save the workbook.
                workbook.Save(filePath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                // Wrap the exception to provide context.
                throw new ApplicationException("Failed to toggle texture tiling.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Adjust the file path and shape name as needed.
                string filePath = @"C:\Docs\Sample.xlsx";
                string shapeName = "MyShape";

                ShapeTextureTiler.ToggleTextureTiling(filePath, shapeName);
                Console.WriteLine("Texture tiling toggled successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
