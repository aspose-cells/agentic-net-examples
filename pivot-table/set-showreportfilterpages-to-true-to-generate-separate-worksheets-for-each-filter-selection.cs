// Title: Create a separate worksheet for each pivot table report filter using ShowReportFilterPage in Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, locates the first pivot table, and calls ShowReportFilterPage for every page field to produce individual worksheets. | Write a method that enables report‑filter page generation on a pivot table, iterates through its PageFields collection, and saves the workbook so each filter value appears on its own sheet. | Provide a step‑by‑step example that checks for a pivot table, uses PivotTable.ShowReportFilterPage, and outputs a new workbook where each report filter selection is placed on a separate worksheet.
// Common Searches: how to use ShowReportFilterPage in Aspose.Cells to split pivot table filters into separate sheets | c# aspnet generate worksheet per pivot table page field aspose.cells | asp.net create multiple worksheets from pivot table report filter using Aspose.Cells API | example code for ShowReportFilterPages property Aspose.Cells .NET
// Tags: Aspose.Cells pivot report filter sheet generation | C# create worksheet per pivot filter selection | Aspose.Cells workbook save with multiple pivot filter sheets | automate pivot table page fields to separate worksheets .NET | export pivot filter pages using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The sample loads 'source.xlsx', verifies a pivot table exists, iterates over its PageFields, calls ShowReportFilterPage for each field to generate distinct worksheets for every filter selection, and saves the result as 'output.xlsx'.
    public class ShowReportFilterPagesDemo
    {
        public static void Run()
        {
            try
            {
                string inputPath = "source.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook that contains a pivot table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Get the first pivot table on the worksheet
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Generate a separate worksheet for each report filter (page field)
                foreach (PivotField pageField in pivotTable.PageFields)
                {
                    pivotTable.ShowReportFilterPage(pageField);
                }

                // Save the modified workbook; each filter selection now has its own sheet
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShowReportFilterPagesDemo.Run();
        }
    }
}
