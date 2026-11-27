// Title: Export each shape's XML definition from an Excel workbook to separate files using Aspose.Cells for .NET (C#)
// Common Searches: how to extract shape XML from an Excel file using Aspose.Cells C# | save Excel worksheet shapes as individual XML files with Aspose.Cells .NET | Aspose.Cells export shape metadata to XML per worksheet C# example | C# code to write shape definitions to separate XML files from a workbook | generate XML representation of shapes in an Aspose.Cells workbook
// Tags: Aspose.Cells shape definition extraction C# | serialize shape attributes to XML files .NET | iterate worksheet shapes using Aspose.Cells API | sanitize worksheet and shape names for file system | minimal XML representation for Excel shapes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeXmlExporter
{
    // The utility validates command‑line arguments, loads the specified Excel workbook with Aspose.Cells, iterates through each worksheet's Shape collection, builds a minimal XML snippet containing each shape's name and type, sanitizes worksheet and shape names for safe file names, and writes each XML definition to an individual file in a configurable output folder.
    class Program
    {
        /// <param name="args">
        /// args[0] - path to the source Excel file (e.g., "input.xlsx").
        /// args[1] - (optional) path to the output folder where XML files will be written.
        /// </param>
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ShapeXmlExporter <input_excel_path> [output_folder]");
                return;
            }

            string inputPath = args[0];
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            string outputFolder = args.Length > 1 ? args[1] : "ShapeXmlExports";

            try
            {
                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    ShapeCollection shapes = sheet.Shapes;

                    if (shapes == null || shapes.Count == 0)
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" contains no shapes.");
                        continue;
                    }

                    Console.WriteLine($"Processing {shapes.Count} shape(s) in worksheet \"{sheet.Name}\"...");

                    // Export each shape's XML definition (fallback to simple XML if native export not supported)
                    for (int i = 0; i < shapes.Count; i++)
                    {
                        Shape shape = shapes[i];
                        string shapeXml;

                        try
                        {
                            // Aspose.Cells .NET does not expose a direct GetXmlData method.
                            // Construct a minimal XML representation containing key attributes.
                            shapeXml = $"<Shape Name=\"{EscapeXml(shape.Name)}\" Type=\"{shape.Type}\"/>";
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error while creating XML for shape \"{shape.Name}\": {ex.Message}");
                            continue;
                        }

                        // Build a safe file name
                        string safeSheetName = MakeFileSystemSafe(sheet.Name);
                        string safeShapeName = MakeFileSystemSafe(shape.Name);
                        string fileName = $"{safeSheetName}_Shape{i + 1}_{safeShapeName}.xml";
                        string fullPath = Path.Combine(outputFolder, fileName);

                        try
                        {
                            File.WriteAllText(fullPath, shapeXml);
                            Console.WriteLine($"Exported shape \"{shape.Name}\" to \"{fullPath}\".");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to write XML file for shape \"{shape.Name}\": {ex.Message}");
                        }
                    }
                }

                Console.WriteLine("Shape XML export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }

        private static string MakeFileSystemSafe(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name;
        }

        /// <summary>
        /// Escapes special XML characters in a string.
        /// </summary>
        private static string EscapeXml(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }
    }
}
