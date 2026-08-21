// Title: Export Excel to CSV with Aspose.Cells LightCells API (C#) – No Workbook Model Load
// Description: Shows how to convert an .xlsx file to a CSV file using Aspose.Cells ConversionUtility in LightCells mode, which processes the workbook without fully loading the model and keeps memory usage low.
// Keywords: Aspose.Cells | LightCells | CSV export | ConversionUtility | C# | .NET | memory‑efficient conversion | large Excel to CSV | streaming Excel conversion | without loading workbook model
// Common Searches: Aspose.Cells export Excel to CSV without loading workbook | LightCells mode CSV conversion C# | Convert large .xlsx to CSV with low memory usage | ConversionUtility Convert method example | C# stream Excel to CSV Aspose.Cells
// Developer Intent: Convert an Excel workbook to a CSV file using Aspose.Cells LightCells mode so the workbook model is not fully loaded into memory.
// Use Cases: Process massive Excel files in data pipelines while staying within RAM limits. | Generate CSV reports from user‑uploaded spreadsheets in web apps without high memory overhead. | Batch‑convert multiple workbooks to CSV in a background service with minimal resource consumption.
// AI Prompts: Provide C# code that uses Aspose.Cells ConversionUtility to export an .xlsx file to CSV in LightCells mode. | Explain how to set a custom delimiter and encoding when converting Excel to CSV with Aspose.Cells LightCells API. | Show an example of handling errors while batch converting Excel files to CSV using ConversionUtility.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Shows how to convert an .xlsx file to a CSV file using Aspose.Cells ConversionUtility in LightCells mode, which processes the workbook without fully loading the model and keeps memory usage low.
class ExportWorkbookToCsv
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Desired CSV output path
        string csvPath = "output.csv";

        // Convert the workbook to CSV.
        // ConversionUtility internally uses LightCells mode, so the workbook model is not fully loaded into memory.
        ConversionUtility.Convert(sourcePath, csvPath);

        Console.WriteLine($"Workbook '{sourcePath}' has been exported to CSV at '{csvPath}'.");
    }
}
