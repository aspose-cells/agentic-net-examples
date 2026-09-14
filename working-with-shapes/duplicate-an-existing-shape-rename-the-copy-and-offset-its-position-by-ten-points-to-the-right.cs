// Title: Duplicate a worksheet shape, rename the copy, and offset it 10 points to the right using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to copy an existing shape on a worksheet, assign a new name to the duplicate, and move it 10 points horizontally. | Show how to call Worksheet.Shapes.AddCopy, set the duplicated shape's Name property, and increase its Left property by 10 points in a .NET workbook.
// Common Searches: C# Aspose.Cells copy shape and change its name on the same worksheet | How to move a duplicated shape 10 points to the right with Aspose.Cells | Worksheet.Shapes.AddCopy example with offset in Aspose.Cells for .NET | Rename copied shape in Aspose.Cells C# code | Shift shape position programmatically using Aspose.Cells Left property
// Tags: Worksheet.Shapes.AddCopy shape duplication | set duplicated shape Name property | modify shape Left offset points | Aspose.Cells shape copy and reposition | C# workbook shape manipulation example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Loads a workbook, locates a shape named "OriginalShape", creates a copy on the same worksheet using Worksheet.Shapes.AddCopy, renames the copy by appending "_Copy", shifts it 10 points to the right by increasing its Left property, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Load existing workbook or create a new one if the file is missing.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // Access the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the shape named "OriginalShape".
                Shape originalShape = worksheet.Shapes["OriginalShape"];
                if (originalShape == null)
                {
                    Console.WriteLine("Shape \"OriginalShape\" not found. No duplication performed.");
                }
                else
                {
                    try
                    {
                        // Duplicate the shape on the same worksheet using AddCopy overload.
                        Shape duplicatedShape = worksheet.Shapes.AddCopy(
                            originalShape,
                            originalShape.UpperLeftRow,
                            originalShape.UpperLeftColumn,
                            originalShape.LowerRightRow,
                            originalShape.LowerRightColumn);

                        // Rename the duplicated shape.
                        duplicatedShape.Name = originalShape.Name + "_Copy";

                        // Offset the duplicated shape 10 points to the right.
                        duplicatedShape.Left = originalShape.Left + 10;
                    }
                    catch (Exception dupEx)
                    {
                        Console.WriteLine($"Failed to duplicate shape: {dupEx.Message}");
                    }
                }

                // Ensure the directory for the output file exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
