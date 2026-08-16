// Title: Highlight Intersection of Two Named Ranges and Save as XLSX with Aspose.Cells for .NET (C#)
// Description: C# example that creates a workbook, defines two named ranges (A1:C3 as FirstRange and B2:D4 as SecondRange), obtains their intersected area, applies a solid yellow background style, and saves the result as an XLSX file using Aspose.Cells.
// Keywords: Aspose.Cells C# | named range intersection | highlight intersected range | apply background color Aspose.Cells | range.Intersect Aspose.Cells | Excel cell formatting .NET | save workbook as XLSX | Excel overlapping ranges | C# Excel styling | Aspose.Cells range style
// Common Searches: Aspose.Cells intersect two named ranges C# | How to color intersecting cells with Aspose.Cells | Set background color for range intersection Aspose.Cells .NET | Get intersected range in Aspose.Cells | Save highlighted Excel file using Aspose.Cells
// Developer Intent: Color the cells where two named ranges overlap and export the workbook as an XLSX file.
// Use Cases: Emphasize overlapping sections in financial reports. | Mark cells that belong to multiple categories in dashboard data. | Provide visual cues for intersecting zones in inventory spreadsheets. | Assist auditors by highlighting shared data areas. | Automate conditional formatting for dynamic range intersections.
// AI Prompts: Write C# code using Aspose.Cells to find the intersection of two named ranges and fill it with a red background. | Show how to apply a gradient fill to the intersected range instead of a solid color. | Demonstrate saving the highlighted workbook to a MemoryStream and returning it as a byte array. | Explain how to retrieve the address of the intersected range and log it. | Provide a version that creates the named ranges at runtime based on user input.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace IntersectNamedRangesDemo
{
    // C# example that creates a workbook, defines two named ranges (A1:C3 as FirstRange and B2:D4 as SecondRange), obtains their intersected area, applies a solid yellow background style, and saves the result as an XLSX file using Aspose.Cells.
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
                worksheet.Cells["A1"].PutValue("A1");
                worksheet.Cells["B1"].PutValue("B1");
                worksheet.Cells["C1"].PutValue("C1");
                worksheet.Cells["A2"].PutValue("A2");
                worksheet.Cells["B2"].PutValue("B2");
                worksheet.Cells["C2"].PutValue("C2");
                worksheet.Cells["D2"].PutValue("D2");
                worksheet.Cells["A3"].PutValue("A3");
                worksheet.Cells["B3"].PutValue("B3");
                worksheet.Cells["C3"].PutValue("C3");
                worksheet.Cells["D3"].PutValue("D3");
                worksheet.Cells["B4"].PutValue("B4");
                worksheet.Cells["C4"].PutValue("C4");
                worksheet.Cells["D4"].PutValue("D4");

                // Define two ranges and assign names to them
                AsposeRange firstRange = worksheet.Cells.CreateRange("A1:C3");
                firstRange.Name = "FirstRange";

                AsposeRange secondRange = worksheet.Cells.CreateRange("B2:D4");
                secondRange.Name = "SecondRange";

                // Get the intersected area of the two ranges
                AsposeRange intersected = firstRange.Intersect(secondRange);

                if (intersected != null)
                {
                    // Create a style with a solid background color
                    Style highlightStyle = workbook.CreateStyle();
                    highlightStyle.Pattern = BackgroundType.Solid;
                    highlightStyle.ForegroundColor = Color.Yellow; // background color

                    // Apply the style to the intersected range
                    intersected.SetStyle(highlightStyle);
                }

                // Save the workbook as XLSX
                workbook.Save("IntersectedRanges.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
