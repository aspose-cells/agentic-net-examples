// Title: Load an existing Excel workbook and insert a line sparkline with markers using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that verifies the presence of an input .xlsx file, loads it with Aspose.Cells, adds a line sparkline to cell C2 sourced from B2:B6, enables ShowMarkers, and saves the result to a new file. | Show how to import the Aspose.Cells.Charts namespace and use SparklineGroups.Add to create a SparklineGroup of type Line with a specified data range in Aspose.Cells. | Provide a robust error‑handling pattern for handling a missing workbook file when adding a sparkline with Aspose.Cells in C#.
// Common Searches: C# Aspose.Cells add line sparkline to existing worksheet example | How to enable markers for a sparkline using Aspose.Cells .NET | Using SparklineGroups.Add with cell range B2:B6 in Aspose.Cells C# | Load workbook and create sparkline group Aspose.Cells code sample | Aspose.Cells sparkline line type C# tutorial
// Tags: add line sparkline Aspose.Cells C# | sparklinegroup add method Aspose.Cells | showmarkers property sparkline Aspose.Cells | load workbook add sparkline Aspose.Cells | import Aspose.Cells.Charts namespace

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program checks for Sample.xlsx, loads it with Aspose.Cells, creates a line sparkline in cell C2 using the B2:B6 range, turns on markers, and saves the modified workbook as Output.xlsx, with error handling for missing input files.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "Sample.xlsx";
            string outputPath = "Output.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the data range for the sparkline (B2:B6)
            CellArea dataRange = CellArea.CreateCellArea("B2", "B6");

            // Add a line sparkline group: source data B2:B6, placed at C2
            // The Add method returns the index of the created group
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "C2", false, dataRange);
            SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

            // Optional: configure sparkline display options (example: show markers)
            sparklineGroup.ShowMarkers = true;

            // Save the workbook with the added sparkline
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
