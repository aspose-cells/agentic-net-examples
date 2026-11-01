// Title: Apply the built‑in Style20 chart style to an existing or newly created chart in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook with Aspose.Cells, locate the first worksheet, and set the Style property of its first chart to the built‑in style index 20. | When a worksheet has no charts, create a column chart, assign Style20, and save the workbook to a new file using C#. | Add robust error handling while loading the workbook, applying a predefined chart style, and writing the styled workbook back to disk.
// Common Searches: how to set chart style 20 in Aspose.Cells C# example | apply predefined chart style to existing chart using Aspose.Cells for .NET | create a column chart and assign built‑in style before saving workbook Aspose.Cells | Aspose.Cells chart style index 20 error handling C#
// Tags: Aspose.Cells set chart style index | apply built‑in chart style .NET | create column chart if none exists Aspose.Cells | load workbook and style chart C# | save workbook with styled chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The C# program loads 'input.xlsx', accesses the first worksheet, uses the existing chart or creates a new column chart, applies the built‑in Style20 (index 20) to the chart, and saves the modified workbook as 'output.xlsx' with comprehensive error handling for missing files and style assignment issues.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the existing workbook (lifecycle: load)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve an existing chart or create a new one if none exist
        Chart chart;
        if (sheet.Charts.Count > 0)
        {
            // Use the first chart in the worksheet
            chart = sheet.Charts[0];
        }
        else
        {
            // Create a new Column chart as a placeholder (lifecycle: create)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            chart = sheet.Charts[chartIndex];
        }

        // Apply a predefined chart style using the integer style index.
        try
        {
            // Style index 20 corresponds to a built‑in chart style.
            chart.Style = 20;
        }
        catch (Exception ex)
        {
            // Log but continue if style assignment fails.
            Console.WriteLine($"Warning: Unable to set chart style. {ex.Message}");
        }

        try
        {
            // Save the workbook with the applied style (lifecycle: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
