// Title: Add major tick marks to the X‑axis of an Excel chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set major tick marks on chart category axis | How to display cross tick marks on Excel chart X axis with Aspose.Cells | Enable major tick marks for X‑axis in a column chart using Aspose.Cells .NET | C# Aspose.Cells chart formatting add major tick marks to X axis | Set CategoryAxis.MajorTickMark property in Aspose.Cells example
// Tags: Aspose.Cells CategoryAxis MajorTickMark | C# Excel chart tick mark customization | Aspose.Cells set X axis tick marks | Excel chart major tick marks .NET | Aspose.Cells chart axis formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads or creates an Excel workbook, accesses the first worksheet, obtains an existing chart or adds a new column chart, sets the CategoryAxis.MajorTickMark property to TickMarkType.Cross (or TickMarkType.Outside) to enable major tick marks on the X‑axis, and saves the workbook.
public class Program
{
    public static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file does not exist
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get an existing chart or create a new column chart as a placeholder
            Chart chart;
            if (sheet.Charts.Count > 0)
            {
                chart = sheet.Charts[0];
            }
            else
            {
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                chart = sheet.Charts[chartIndex];
            }

            // Enable major tick marks on the X‑axis (category axis)
            chart.CategoryAxis.MajorTickMark = TickMarkType.Cross; // or TickMarkType.Outside

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
