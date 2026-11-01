// Title: Delete all shapes from an Excel worksheet with Aspose.Cells for .NET and then export to CSV
// AI Prompts: Generate C# code that iterates the Shapes collection in reverse, removes each shape, and saves the workbook as a CSV file using Aspose.Cells. | Provide a method that accepts an .xlsx file path, clears all drawings from the first worksheet, and writes the cleaned data to a .csv file with Aspose.Cells. | Show how to programmatically delete pictures, charts, and other shapes from a worksheet before exporting the sheet to CSV in .NET.
// Common Searches: how to programmatically delete all drawings from an Excel worksheet using Aspose.Cells .NET | export Excel to CSV without embedded shapes Aspose.Cells example | remove charts and pictures from worksheet before CSV conversion C# Aspose.Cells | best way to clear Shapes collection in Aspose.Cells prior to saving as CSV
// Tags: Aspose.Cells remove worksheet shapes | reverse iteration shape deletion Aspose.Cells | export cleaned worksheet to CSV Aspose.Cells | clear drawings before CSV conversion .NET | delete all shapes from Excel worksheet C#

using System;
using Aspose.Cells;

// Loads an Excel workbook, iterates the Shapes collection backward to delete every shape from the first worksheet, and saves the cleaned worksheet as a CSV file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify by name/index)
        Worksheet worksheet = workbook.Worksheets[0];

        // Remove all shapes from the worksheet
        // Iterate backwards because removing items shifts the collection indices
        for (int i = worksheet.Shapes.Count - 1; i >= 0; i--)
        {
            worksheet.Shapes.RemoveAt(i);
        }

        // Export the cleaned worksheet to CSV format
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
