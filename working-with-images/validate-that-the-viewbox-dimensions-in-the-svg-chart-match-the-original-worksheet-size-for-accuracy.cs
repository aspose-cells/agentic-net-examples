// Title: Check that an exported SVG viewBox matches the worksheet pixel dimensions in Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, computes the total pixel width and height of the used range, saves the sheet as SVG, reads the root <svg> viewBox attribute, and verifies the width and height against the calculated pixel values using a small tolerance. | Create a reusable C# method that accepts a Worksheet, calculates its pixel size via GetColumnWidthPixel and GetRowHeightPixel, extracts the viewBox from an SVG stream produced by Aspose.Cells, and returns whether the dimensions are within an acceptable tolerance.
// Common Searches: Aspose.Cells how to compare worksheet pixel size to SVG viewBox after export | C# validate SVG viewBox dimensions against Excel sheet dimensions | calculate total column and row pixel dimensions in Aspose.Cells for viewBox verification
// Tags: Aspose.Cells export worksheet to SVG | calculate worksheet pixel dimensions Aspose.Cells | validate SVG viewBox against Excel sheet | GetColumnWidthPixel GetRowHeightPixel usage | SVG viewBox verification C#

using System;
using System.Globalization;
using System.IO;
using System.Xml;
using Aspose.Cells;

// The example loads an Excel workbook, determines the used range of the first worksheet, sums column widths and row heights in pixels, exports the worksheet (including charts) to SVG, parses the root <svg> element's viewBox attribute, extracts its width and height, and compares these values to the calculated worksheet pixel dimensions within a 0.5‑pixel tolerance, reporting success or any mismatches.
class SvgViewBoxValidator
{
    static void Main()
    {
        try
        {
            // Verify that the input workbook exists
            const string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet (use Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            // Calculate total width in pixels by summing column widths
            double totalWidthPx = 0;
            int firstCol = usedRange.FirstColumn;
            int lastCol = firstCol + usedRange.ColumnCount - 1;
            for (int col = firstCol; col <= lastCol; col++)
            {
                totalWidthPx += sheet.Cells.GetColumnWidthPixel(col);
            }

            // Calculate total height in pixels by summing row heights
            double totalHeightPx = 0;
            int firstRow = usedRange.FirstRow;
            int lastRow = firstRow + usedRange.RowCount - 1;
            for (int row = firstRow; row <= lastRow; row++)
            {
                totalHeightPx += sheet.Cells.GetRowHeightPixel(row);
            }

            // Export the worksheet (including charts) to SVG using a memory stream
            string svgContent;
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Svg);
                ms.Position = 0;
                using (StreamReader reader = new StreamReader(ms))
                {
                    svgContent = reader.ReadToEnd();
                }
            }

            if (string.IsNullOrEmpty(svgContent))
            {
                Console.WriteLine("Failed to generate SVG content.");
                return;
            }

            // Load the SVG XML
            XmlDocument svgDoc = new XmlDocument();
            svgDoc.LoadXml(svgContent);

            // Retrieve the viewBox attribute from the root <svg> element
            XmlElement root = svgDoc.DocumentElement;
            if (root == null)
            {
                Console.WriteLine("Invalid SVG content: missing root element.");
                return;
            }

            string viewBox = root.GetAttribute("viewBox");
            if (string.IsNullOrEmpty(viewBox))
            {
                Console.WriteLine("No viewBox attribute found in the SVG.");
                return;
            }

            // viewBox format: "minX minY width height"
            string[] parts = viewBox.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 4)
            {
                Console.WriteLine("Invalid viewBox format.");
                return;
            }

            // Parse width and height from viewBox using invariant culture
            if (!double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double viewBoxWidth) ||
                !double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out double viewBoxHeight))
            {
                Console.WriteLine("Unable to parse viewBox dimensions.");
                return;
            }

            // Compare the viewBox dimensions with the worksheet pixel dimensions
            const double tolerance = 0.5; // allow a small rounding tolerance

            bool widthMatches = Math.Abs(viewBoxWidth - totalWidthPx) <= tolerance;
            bool heightMatches = Math.Abs(viewBoxHeight - totalHeightPx) <= tolerance;

            if (widthMatches && heightMatches)
            {
                Console.WriteLine("Success: SVG viewBox dimensions match the worksheet size.");
            }
            else
            {
                Console.WriteLine("Mismatch detected:");
                if (!widthMatches)
                    Console.WriteLine($" - Width: viewBox={viewBoxWidth}, worksheet={totalWidthPx}");
                if (!heightMatches)
                    Console.WriteLine($" - Height: viewBox={viewBoxHeight}, worksheet={totalHeightPx}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
