// Title: Insert a line sparkline in cell P5 from range B2:B10 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new workbook, defines a SparklineGroup of type Line for the range B2:B10, places it in cell P5, enables markers, first‑point and last‑point visibility, and saves the file as an .xlsx. | Show how to obtain the index of a newly added SparklineGroup, adjust its ShowMarkers, ShowFirstPoint, and ShowLastPoint properties, and persist the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# add line sparkline to a single cell | Create sparkline in Excel using Aspose.Cells from a vertical range | C# example for inserting sparkline in cell P5 with Aspose.Cells | How to enable markers, first point, and last point on an Aspose.Cells sparkline | Saving workbook after adding sparkline with Aspose.Cells .NET
// Tags: Aspose.Cells sparkline insertion | C# SparklineGroup appearance settings | Insert sparkline into specific worksheet cell | Generate Excel sparkline from data range B2:B10 | Aspose.Cells workbook save with sparkline

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, adds a line sparkline that references cells B2:B10 to cell P5, turns on markers as well as first‑point and last‑point highlights, and saves the result as SparklineExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the data range for the sparkline
            string dataRange = "B2:B10";

            // Define the location cell for the sparkline (P5)
            // Create a CellArea that represents a single cell (P5)
            CellArea location = CellArea.CreateCellArea("P5", "P5");

            // Add a line sparkline group (isVertical = false)
            // The Add method returns the index of the created group
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, dataRange, false, location);
            SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

            // Optional: customize sparkline appearance
            sparklineGroup.ShowMarkers = true;
            sparklineGroup.ShowFirstPoint = true;
            sparklineGroup.ShowLastPoint = true;

            // Define output file path
            string outputPath = "SparklineExample.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
