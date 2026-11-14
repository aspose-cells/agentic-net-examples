// Title: Measure smart marker processing time in C# with Aspose.Cells by logging start and end timestamps
// Common Searches: how to log start and end time of Aspose.Cells smart marker processing in C# | C# code to measure execution duration of WorkbookDesigner.Process for smart markers | track smart marker processing performance using Aspose.Cells .NET | log timestamps before and after smart marker replacement with Aspose.Cells | measure total time taken by smart markers in a .xlsx workbook using C#
// Tags: smart marker processing timing Aspose.Cells | WorkbookDesigner execution duration logging | C# Aspose.Cells performance tracking | log smart marker timestamps .NET | measure smart marker execution time

using System;
using System.IO;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerTimingExample
{
    // Simple timing helper class (no Aspose callback interface required)
    // The example creates a helper class that records the start time when instantiated, runs WorkbookDesigner.Process to fill smart markers, then logs the end time and total elapsed time. It demonstrates loading or creating a workbook, supplying a DataSet as the data source, processing smart markers, saving the result, and outputting timing information to the console.
    public class SmartMarkerTimingHelper
    {
        private readonly DateTime _startTime;

        public SmartMarkerTimingHelper()
        {
            _startTime = DateTime.Now;
            Console.WriteLine($"Smart marker processing started at {_startTime:O}");
        }

        public void LogCompletion()
        {
            DateTime endTime = DateTime.Now;
            Console.WriteLine($"Smart marker processing ended at {endTime:O}");
            Console.WriteLine($"Total execution time: {endTime - _startTime}");
        }
    }

    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a minimal workbook if it does not.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a placeholder workbook.");
                    var placeholderWorkbook = new Workbook();
                    var ws = placeholderWorkbook.Worksheets[0];
                    ws.Cells["A1"].PutValue("Smart Marker Example");
                    placeholderWorkbook.Save(inputPath);
                }

                // Load the workbook that contains smart markers
                var workbook = new Workbook(inputPath);

                // Prepare the data source for smart markers using a DataSet
                var dataTable = new DataTable("Employees");
                dataTable.Columns.Add("Name", typeof(string));
                dataTable.Columns.Add("Age", typeof(int));
                dataTable.Rows.Add("John Doe", 30);
                dataTable.Rows.Add("Jane Smith", 28);

                var dataSet = new DataSet();
                dataSet.Tables.Add(dataTable);

                // Set up the WorkbookDesigner with the data source
                var designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSet);

                // Timing helper
                var timingHelper = new SmartMarkerTimingHelper();

                // Process smart markers (no custom callback needed)
                designer.Process();

                // Log completion time
                timingHelper.LogCompletion();

                // Ensure the output directory exists
                var outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the processed workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processed workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
