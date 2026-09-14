// Title: Load an XLSX workbook from a file path and retrieve the first chart with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to open a .xlsx file from a specified path, verifies that the first worksheet contains charts, and returns the first Chart object. | Write a method that loads a workbook, checks for the existence of the file, handles missing‑file errors, and prints the type of the first chart in the first worksheet. | Create a reusable function that accepts a file path, loads the workbook with Aspose.Cells, and extracts properties of the first chart for further processing.
// Common Searches: Aspose.Cells C# load workbook from file path and get first chart in worksheet | How to read chart type from an XLSX file using Aspose.Cells .NET | C# example for checking chart collection count before accessing chart with Aspose.Cells | Retrieve chart objects from a worksheet using Aspose.Cells in a console application | Handle FileNotFoundException when opening Excel file with Aspose.Cells C#
// Tags: load workbook from file path Aspose.Cells | access first worksheet chart collection Aspose.Cells | retrieve first chart object C# | read chart type from XLSX Aspose.Cells | handle missing Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartDemo
{
    // Demonstrates loading an existing XLSX file with Aspose.Cells, accessing the first worksheet, verifying that it contains at least one chart, obtaining the first Chart object, and outputting its type while handling missing‑file scenarios.
    public class LoadAndGetFirstChart
    {
        public static void Run()
        {
            // Path to the existing XLSX file
            string filePath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook using the string constructor
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet in the workbook
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                    return;
                }

                // Obtain the first chart object from the worksheet's chart collection
                Chart firstChart = worksheet.Charts[0];

                // Demonstrate accessing a property of the chart (e.g., its type)
                Console.WriteLine($"First chart type: {firstChart.Type}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            LoadAndGetFirstChart.Run();
        }
    }
}
