// Title: Highlight intersecting cells of two named ranges with yellow fill using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# find overlapping cells of two named ranges | how to calculate intersection of named ranges in an Excel workbook with Aspose.Cells | C# code to apply yellow background to intersected range using Aspose.Cells | retrieve worksheet from a named range reference in Aspose.Cells | Aspose.Cells create style with solid fill and apply to a range
// Tags: named range intersection Aspose.Cells | apply yellow fill style Aspose.Cells | CellArea overlap calculation C# | retrieve worksheet from Name object Aspose.Cells | create solid background style Aspose.Cells | Excel range intersection C#

using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;
using System;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook, obtains two named ranges, parses their references into CellArea objects, computes the intersecting rectangle, creates a solid yellow style, applies the style to the intersected cells on the appropriate worksheet, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the input file
            var workbook = new Workbook(inputPath);

            // Retrieve the two named ranges from the workbook's Names collection
            var namedRange1 = workbook.Worksheets.Names["Range1"];
            var namedRange2 = workbook.Worksheets.Names["Range2"];

            if (namedRange1 == null || namedRange2 == null)
            {
                Console.WriteLine("One or both named ranges were not found.");
                return;
            }

            // Convert the reference strings of the named ranges into CellArea structures
            CellArea area1 = GetCellAreaFromName(namedRange1);
            CellArea area2 = GetCellAreaFromName(namedRange2);

            // Compute the intersecting rectangle
            int intersectStartRow = Math.Max(area1.StartRow, area2.StartRow);
            int intersectStartCol = Math.Max(area1.StartColumn, area2.StartColumn);
            int intersectEndRow   = Math.Min(area1.EndRow,   area2.EndRow);
            int intersectEndCol   = Math.Min(area1.EndColumn,area2.EndColumn);

            // Verify that an intersection exists
            if (intersectStartRow <= intersectEndRow && intersectStartCol <= intersectEndCol)
            {
                // Get the worksheet that contains the first named range
                Worksheet worksheet = GetWorksheetFromName(namedRange1, workbook);

                // Create a style with solid yellow fill
                Style yellowStyle = workbook.CreateStyle();
                yellowStyle.ForegroundColor = Color.Yellow;
                yellowStyle.Pattern = BackgroundType.Solid;

                StyleFlag styleFlag = new StyleFlag { CellShading = true };

                // Define the intersecting range and apply the style
                int rows = intersectEndRow - intersectStartRow + 1;
                int cols = intersectEndCol - intersectStartCol + 1;
                AsposeRange intersectRange = worksheet.Cells.CreateRange(intersectStartRow, intersectStartCol, rows, cols);
                intersectRange.ApplyStyle(yellowStyle, styleFlag);
            }
            else
            {
                Console.WriteLine("The named ranges do not intersect.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Parses a Name.RefersTo string and returns a CellArea representing the range.
    private static CellArea GetCellAreaFromName(Name name)
    {
        // RefersTo format: SheetName!$A$1:$C$5 (or a single cell)
        string refersTo = name.RefersTo;
        int exclPos = refersTo.IndexOf('!');
        string rangePart = exclPos >= 0 ? refersTo.Substring(exclPos + 1) : refersTo;

        // Remove any leading '$' characters
        rangePart = rangePart.Replace("$", string.Empty);

        // Split start and end cells
        string[] cells = rangePart.Split(':');
        if (cells.Length == 1)
        {
            // Single cell range
            return CellArea.CreateCellArea(cells[0], cells[0]);
        }
        else
        {
            return CellArea.CreateCellArea(cells[0], cells[1]);
        }
    }

    // Retrieves the worksheet referenced by a Name object.
    private static Worksheet GetWorksheetFromName(Name name, Workbook workbook)
    {
        string refersTo = name.RefersTo;
        int exclPos = refersTo.IndexOf('!');
        if (exclPos > 0)
        {
            string sheetName = refersTo.Substring(0, exclPos);
            // Trim possible surrounding single quotes
            sheetName = sheetName.Trim('\'');
            return workbook.Worksheets[sheetName];
        }
        // Fallback to the first worksheet if the sheet name cannot be determined
        return workbook.Worksheets[0];
    }
}
