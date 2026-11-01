// Title: Enumerate child shapes of a GroupShape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, finds the first GroupShape on the first worksheet, and prints each child shape's name, type, left, top, width, and height. | Show a .NET example that uses reflection to retrieve the GroupShapeItems collection from a GroupShape for compatibility with older Aspose.Cells releases. | Create a script that iterates through all shapes inside a GroupShape, outputs their properties, and saves the workbook to a new file.
// Common Searches: Aspose.Cells C# enumerate shapes inside a GroupShape on a worksheet | how to get child shapes of a grouped shape using Aspose.Cells .NET | access GroupShapeItems via reflection Aspose.Cells example | list individual components of a GroupShape in Excel with Aspose.Cells | C# code to print position and size of shapes within a GroupShape
// Tags: enumerate group shape child shapes Aspose.Cells | reflection access GroupShapeItems .NET | list shape properties name type position size Aspose.Cells | load workbook find first GroupShape C# | save workbook after shape enumeration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, locates the first GroupShape on the first worksheet, uses reflection to obtain the GroupShapeItems collection when available, iterates through each child shape printing its name, type, position, and dimensions, and finally saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a placeholder if missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
                workbook.Worksheets.Add("Sheet1");
                workbook.Save(inputPath);
            }

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Get all shapes on the worksheet.
            ShapeCollection shapes = worksheet.Shapes;

            // Find the first GroupShape, if any.
            GroupShape groupShape = null;
            foreach (Shape shape in shapes)
            {
                if (shape is GroupShape)
                {
                    groupShape = (GroupShape)shape;
                    break;
                }
            }

            if (groupShape == null)
            {
                Console.WriteLine("No GroupShape found on the worksheet.");
            }
            else
            {
                Console.WriteLine($"GroupShape found: Name = {groupShape.Name}, Type = {groupShape.Type}");

                // Attempt to enumerate child shapes via reflection (covers versions without GroupShapeItems).
                var propInfo = groupShape.GetType().GetProperty("GroupShapeItems");
                if (propInfo != null)
                {
                    var childCollection = propInfo.GetValue(groupShape) as ShapeCollection;
                    if (childCollection != null)
                    {
                        foreach (Shape childShape in childCollection)
                        {
                            Console.WriteLine($"Child Shape - Name: {childShape.Name}, Type: {childShape.Type}");
                            Console.WriteLine($"  Position: Left={childShape.Left}, Top={childShape.Top}");
                            Console.WriteLine($"  Size: Width={childShape.Width}, Height={childShape.Height}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("GroupShapeItems property exists but returned null.");
                    }
                }
                else
                {
                    Console.WriteLine("The current Aspose.Cells version does not expose GroupShapeItems.");
                }
            }

            // Save the workbook (optional if modifications were made).
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
