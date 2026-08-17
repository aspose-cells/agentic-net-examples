// Title: Highlight the Intersection of Two Named Ranges and Export as ODS with Aspose.Cells for .NET
// Description: This example creates a workbook, defines two named ranges (A1:C3 and B2:D4), retrieves their overlapping cells, applies a solid yellow fill to the intersected area, and saves the result in OpenDocument Spreadsheet (ODS) format using Aspose.Cells for .NET.
// Keywords: Aspose.Cells .NET intersect ranges | apply background color to cells | named range intersection | save workbook as ODS | C# spreadsheet styling | OpenDocument format Aspose | range styling Aspose.Cells
// Common Searches: Aspose.Cells highlight intersecting cells C# | How to color overlap of two named ranges in .NET | Save styled spreadsheet as ODS using Aspose | C# code for range intersection and fill | Aspose.Cells set background for intersected range
// Developer Intent: The developer needs to programmatically identify the overlapping cells of two named ranges, apply a visual highlight, and generate an ODS file for cross‑platform use.
// Use Cases: Mark common data points when merging reports from different departments. | Visually differentiate cells that belong to multiple categories for quick analysis. | Produce ODS files with highlighted intersections for downstream processing in LibreOffice or other OpenDocument tools.
// AI Prompts: Write C# code with Aspose.Cells that finds the intersection of two named ranges, fills it with red, and saves the workbook as ODS. | Explain step‑by‑step how to create named ranges, get their intersected area, apply a solid background style, and export to ODS in Aspose.Cells for .NET. | Provide enhanced error handling for null intersections and show how to save the same workbook in both ODS and XLSX formats.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsIntersectionDemo
{
    // This example creates a workbook, defines two named ranges (A1:C3 and B2:D4), retrieves their overlapping cells, applies a solid yellow fill to the intersected area, and saves the result in OpenDocument Spreadsheet (ODS) format using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate some sample data
                worksheet.Cells["A1"].PutValue("Data A1");
                worksheet.Cells["B2"].PutValue("Data B2");
                worksheet.Cells["C3"].PutValue("Data C3");
                worksheet.Cells["D4"].PutValue("Data D4");

                // Define two named ranges
                AsposeRange range1 = worksheet.Cells.CreateRange("A1:C3");
                range1.Name = "FirstRange";

                AsposeRange range2 = worksheet.Cells.CreateRange("B2:D4");
                range2.Name = "SecondRange";

                // Get the intersected area of the two ranges
                AsposeRange intersected = range1.Intersect(range2);
                if (intersected != null)
                {
                    // Create a style with a solid background color
                    Style style = workbook.CreateStyle();
                    style.Pattern = BackgroundType.Solid;
                    style.ForegroundColor = Color.Yellow;

                    // Apply the style to the intersected range
                    intersected.SetStyle(style);
                }

                // Save the workbook as ODS
                workbook.Save("IntersectedRanges.ods", SaveFormat.Ods);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
