// Title: C# utility to extract all embedded SVG images from an Excel workbook using Aspose.Cells and save them to a folder
// AI Prompts: Write a C# console program that opens an Excel file with Aspose.Cells, walks through each worksheet, and exports any SVG picture or shape to a user‑specified output directory. | Implement reflection in C# to read the ImageFormatType or raw ImageData of Aspose.Cells Picture and Shape objects, determine whether the content is SVG, and then write the SVG bytes to disk. | Create logic that assigns sequential filenames (e.g., svg_0.svg, svg_1.svg) to each extracted SVG and prints the total count of SVG resources after processing the workbook.
// Common Searches: how to export embedded svg images from an Excel file using Aspose.Cells in C# | C# code to extract svg pictures from worksheets with Aspose.Cells | save svg resources from an Aspose.Cells workbook to a folder | detect svg shape in Aspose.Cells and write to file
// Tags: Aspose.Cells SVG extraction | C# workbook image export | picture and shape iteration Aspose.Cells | reflection based image type detection | sequential naming of exported SVGs

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// A C# console utility that loads an Excel workbook via Aspose.Cells, iterates through each worksheet's pictures and shapes, uses reflection to identify SVG content (by ImageFormatType or byte signature), and saves each SVG as a sequentially named .svg file in a specified output folder, reporting the total number extracted.
class SvgExtractor
{
    static void Main(string[] args)
    {
        // Expect two arguments: path to the workbook and output folder
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: SvgExtractor <workbookPath> <outputFolder>");
            return;
        }

        string workbookPath = args[0];
        string outputFolder = args[1];

        // Verify that the workbook file exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file not found at \"{workbookPath}\".");
            return;
        }

        // Ensure the output directory exists
        try
        {
            Directory.CreateDirectory(outputFolder);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to create output directory: {ex.Message}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(workbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        int svgIndex = 0;

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // ----- Extract SVGs stored as pictures -----
            foreach (Picture picture in sheet.Pictures)
            {
                try
                {
                    // Use reflection to safely access ImageFormatType (may not exist in older versions)
                    PropertyInfo formatProp = picture.GetType().GetProperty("ImageFormatType");
                    bool isSvg = false;

                    if (formatProp != null)
                    {
                        object formatValue = formatProp.GetValue(picture);
                        // Compare with the enum name "Svg" if possible
                        if (formatValue != null && formatValue.ToString().Equals("Svg", StringComparison.OrdinalIgnoreCase))
                            isSvg = true;
                    }

                    // Fallback: inspect raw bytes for SVG signature if format info unavailable
                    PropertyInfo dataProp = picture.GetType().GetProperty("ImageData");
                    if (dataProp != null)
                    {
                        object imgDataObj = dataProp.GetValue(picture);
                        MethodInfo toByteArray = imgDataObj?.GetType().GetMethod("ToByteArray");
                        byte[] bytes = toByteArray?.Invoke(imgDataObj, null) as byte[];

                        if (bytes != null && bytes.Length > 0)
                        {
                            if (!isSvg)
                            {
                                // Simple check for SVG content
                                string header = System.Text.Encoding.UTF8.GetString(bytes, 0, Math.Min(100, bytes.Length));
                                isSvg = header.Contains("<svg", StringComparison.OrdinalIgnoreCase);
                            }

                            if (isSvg)
                            {
                                string fileName = $"svg_{svgIndex++}.svg";
                                string filePath = Path.Combine(outputFolder, fileName);
                                File.WriteAllBytes(filePath, bytes);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing picture in sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // ----- Extract SVGs stored as shapes (if supported) -----
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Attempt to retrieve image data via reflection
                    PropertyInfo dataProp = shape.GetType().GetProperty("ImageData");
                    if (dataProp != null)
                    {
                        object imgDataObj = dataProp.GetValue(shape);
                        MethodInfo toByteArray = imgDataObj?.GetType().GetMethod("ToByteArray");
                        byte[] bytes = toByteArray?.Invoke(imgDataObj, null) as byte[];

                        if (bytes != null && bytes.Length > 0)
                        {
                            // Simple SVG detection
                            string header = System.Text.Encoding.UTF8.GetString(bytes, 0, Math.Min(100, bytes.Length));
                            if (header.Contains("<svg", StringComparison.OrdinalIgnoreCase))
                            {
                                string fileName = $"svg_{svgIndex++}.svg";
                                string filePath = Path.Combine(outputFolder, fileName);
                                File.WriteAllBytes(filePath, bytes);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing shape in sheet \"{sheet.Name}\": {ex.Message}");
                }
            }
        }

        Console.WriteLine($"Extraction complete. {svgIndex} SVG resource(s) saved to \"{outputFolder}\".");
    }
}
