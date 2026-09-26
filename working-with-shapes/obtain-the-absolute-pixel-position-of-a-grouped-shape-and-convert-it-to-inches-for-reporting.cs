// Title: Determine absolute pixel coordinates of a child shape inside a GroupShape and convert them to inches with Aspose.Cells for .NET
// AI Prompts: Locate a GroupShape by its name, retrieve its first child shape, determine the child’s worksheet row and column indices, then translate that location into pixels and inches using the default 96 DPI. | Write C# code that uses Aspose.Cells to obtain a grouped shape’s offsets, add them to a child shape’s offsets, approximate pixel values from row height (points) and column width (character units), and output the final position in inches.
// Common Searches: how to get pixel position of a shape inside a group in Aspose.Cells C# | convert Excel shape coordinates to inches using Aspose.Cells .NET | retrieve absolute row and column of a child shape from GroupShape Aspose.Cells | calculate shape location in inches from row height and column width Aspose.Cells
// Tags: Aspose.Cells GroupShape child position calculation | shape coordinates inches conversion Aspose.Cells | pixel conversion from row height points Aspose | column width character to pixel Aspose.Cells | absolute worksheet cell location for grouped shape

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, finds a GroupShape by name, accesses its first child shape via reflection, computes the child’s absolute worksheet row and column by adding the group’s offsets, approximates pixel values from the row height (points) and column width (character units) using a 96 DPI reference, converts those pixel measurements to inches, and prints the results before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the shape collection of the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Locate the group shape by its name (replace with actual name if different)
            const string targetGroupName = "MyGroupShape";
            Shape groupShape = null;
            foreach (Shape shp in shapes)
            {
                if (shp.Name == targetGroupName)
                {
                    groupShape = shp;
                    break;
                }
            }

            if (groupShape == null)
            {
                Console.WriteLine($"Grouped shape \"{targetGroupName}\" not found.");
                return;
            }

            // Cast to GroupShape
            GroupShape group = groupShape as GroupShape;
            if (group == null)
            {
                Console.WriteLine("The identified shape is not a group shape.");
                return;
            }

            // Retrieve child shapes via reflection to support multiple Aspose.Cells versions
            ShapeCollection childShapes = null;
            try
            {
                PropertyInfo prop = group.GetType().GetProperty("GroupShapeCollection");
                if (prop != null)
                {
                    childShapes = prop.GetValue(group) as ShapeCollection;
                }

                // Fallback to a possible older property name
                if (childShapes == null)
                {
                    prop = group.GetType().GetProperty("Shapes");
                    if (prop != null)
                    {
                        childShapes = prop.GetValue(group) as ShapeCollection;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving child shapes: {ex.Message}");
                return;
            }

            if (childShapes == null || childShapes.Count == 0)
            {
                Console.WriteLine("Group shape does not contain any child shapes.");
                return;
            }

            // Choose the first child shape (or select by index/name as needed)
            Shape childShape = childShapes[0];

            // ----- Compute absolute position -----
            int groupTopRow = group.UpperLeftRow;
            int groupLeftColumn = group.UpperLeftColumn;
            int childTopRow = childShape.UpperLeftRow;
            int childLeftColumn = childShape.UpperLeftColumn;

            // Absolute cell position of the child shape on the worksheet
            int absoluteRow = groupTopRow + childTopRow;
            int absoluteColumn = groupLeftColumn + childLeftColumn;

            // Approximate pixel conversion (using default DPI 96)
            const double pixelsPerInch = 96.0;

            // Approximate row height in points and convert to pixels
            double rowHeightPoints = sheet.Cells.Rows[absoluteRow].Height; // points
            double topPixels = rowHeightPoints * pixelsPerInch / 72.0; // points → inches → pixels

            // Approximate column width in pixels (Aspose uses character width; 1 character ≈ 7.5 pixels)
            double columnWidthChars = sheet.Cells.Columns[absoluteColumn].Width;
            double leftPixels = columnWidthChars * 7.5;

            // Convert to inches for reporting
            double topInches = topPixels / pixelsPerInch;
            double leftInches = leftPixels / pixelsPerInch;

            // Output the results
            Console.WriteLine($"Child shape absolute cell position: Row = {absoluteRow}, Column = {absoluteColumn}");
            Console.WriteLine($"Approximate pixel position: Top = {topPixels:F2} px ({topInches:F2} in), Left = {leftPixels:F2} px ({leftInches:F2} in)");

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (if any modifications were made)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
