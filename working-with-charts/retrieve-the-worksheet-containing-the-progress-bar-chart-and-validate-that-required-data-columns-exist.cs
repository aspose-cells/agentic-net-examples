// Title: Find the worksheet that contains a 'Progress Bar' chart and ensure it has Task, Start, and Finish columns using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that loops through all worksheets, detects the chart titled "Progress Bar", and returns the worksheet that hosts it. | Add logic to the Aspose.Cells example to read the first row of the identified worksheet, verify the presence of the headers "Task", "Start", and "Finish", and throw a descriptive exception if any are missing.
// Common Searches: asp.net locate worksheet with specific chart title using Aspose.Cells | c# validate that Excel sheet contains Task, Start, Finish columns before processing | Aspose.Cells find chart named "Progress Bar" and get its parent worksheet | how to throw exception when required columns are missing in Aspose.Cells workbook | check header row for required columns in Excel file with Aspose.Cells C#
// Tags: find worksheet by chart title Aspose.Cells | validate required headers Excel Aspose.Cells | check Task Start Finish columns C# | exception handling missing chart Aspose.Cells | progress bar chart worksheet retrieval .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace Example
{
    // The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet and its charts to locate a chart whose title equals "Progress Bar", captures the containing worksheet, then scans the first row for the required headers "Task", "Start", and "Finish". If any header is absent, it throws an exception; otherwise, the workbook is saved.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Locate the worksheet that contains the chart titled "Progress Bar"
                Worksheet progressWorksheet = null;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    foreach (Chart chart in ws.Charts)
                    {
                        if (chart.Title != null && chart.Title.Text == "Progress Bar")
                        {
                            progressWorksheet = ws;
                            break;
                        }
                    }
                    if (progressWorksheet != null)
                        break;
                }

                // Throw if the chart (and thus the worksheet) cannot be found
                if (progressWorksheet == null)
                    throw new Exception("Progress Bar chart not found in any worksheet.");

                // Validate that required data columns exist in the worksheet
                // Assumption: the first row (index 0) holds column headers
                Cells cells = progressWorksheet.Cells;
                int headerRowIndex = 0;

                // Define the required column names
                string[] requiredColumns = { "Task", "Start", "Finish" };
                bool[] columnFound = new bool[requiredColumns.Length];

                // Scan the header row cells
                foreach (Cell cell in cells.Rows[headerRowIndex])
                {
                    string header = cell.StringValue.Trim();
                    for (int i = 0; i < requiredColumns.Length; i++)
                    {
                        if (header.Equals(requiredColumns[i], StringComparison.OrdinalIgnoreCase))
                            columnFound[i] = true;
                    }
                }

                // If any required column is missing, raise an exception
                for (int i = 0; i < requiredColumns.Length; i++)
                {
                    if (!columnFound[i])
                        throw new Exception($"Required column \"{requiredColumns[i]}\" is missing in worksheet \"{progressWorksheet.Name}\".");
                }

                // Save the workbook after validation if further processing is needed
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook validated and saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
