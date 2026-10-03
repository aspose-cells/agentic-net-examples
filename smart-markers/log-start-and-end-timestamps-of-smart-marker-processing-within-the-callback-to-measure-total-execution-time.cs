// Title: Log start and end timestamps to measure smart marker processing time using Aspose.Cells for .NET
// AI Prompts: Insert DateTime.Now before and after WorkbookDesigner.Process() and output the elapsed seconds in the console. | Wrap the smart marker processing call with a Stopwatch to log execution duration in a C# console application. | Generate code that prints the start time, end time, and total processing time for Aspose.Cells smart markers.
// Common Searches: how to time Aspose.Cells smart marker processing in C# | C# log execution time of WorkbookDesigner.Process method | measure performance of smart markers with Aspose.Cells .NET | log start and end timestamps around smart marker processing Aspose.Cells
// Tags: smart marker processing timing Aspose.Cells | WorkbookDesigner.Process performance measurement | log execution duration C# Aspose.Cells | timestamp logging for Excel smart markers | measure smart marker processing time .NET

using System;
using System.IO;
using System.Data;
using Aspose.Cells;

// // Loads an input workbook, creates a DataTable as a data source, assigns it to a WorkbookDesigner, logs the start and end DateTime around designer.Process(), calculates the total seconds elapsed, and saves the processed workbook to an output file.
public class Program
{
    public static void Main()
    {
        try
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Create a simple data source
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Score");
            dt.Rows.Add("Alice", 85);
            dt.Rows.Add("Bob", 92);
            dt.Rows.Add("Charlie", 78);

            // Set up the designer with the data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);

            // Process smart markers and log timing
            DateTime startTime = DateTime.Now;
            Console.WriteLine($"Smart marker processing started at {startTime:O}");

            designer.Process();

            DateTime endTime = DateTime.Now;
            Console.WriteLine($"Smart marker processing finished at {endTime:O}");
            Console.WriteLine($"Total processing time: {(endTime - startTime).TotalSeconds} seconds");

            // Save the processed workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
