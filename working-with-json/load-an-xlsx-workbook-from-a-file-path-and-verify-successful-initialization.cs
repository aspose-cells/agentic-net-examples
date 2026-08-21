// Title: C# – Load an XLSX Workbook from a File Path and Verify Worksheets with Aspose.Cells
// Description: Demonstrates how to instantiate a Workbook using the `Workbook(string)` constructor, check that the workbook contains at least one worksheet, output the worksheet count, and display the name of the first sheet for confirmation.
// Keywords: Aspose.Cells load workbook C# | Workbook(string) constructor | verify Excel worksheets .NET | check worksheet count Aspose.Cells | load XLSX file programmatically
// Common Searches: how to open an xlsx file with Aspose.Cells in C# | C# verify workbook has sheets after loading | Aspose.Cells check if workbook is empty | load excel file and get first worksheet name C#
// Developer Intent: Open an XLSX file via a file path, ensure the workbook is initialized, and confirm the presence of worksheets.
// Use Cases: Validate user‑uploaded Excel files on a web service before processing data. | Load a template workbook, read its first sheet name, and then populate it with dynamic content. | Log workbook health (sheet count, first sheet name) during automated ETL jobs.
// AI Prompts: Create a C# method that loads an Excel file with Aspose.Cells, throws a custom exception if no worksheets exist, and returns the worksheet count. | Write reusable C# code that accepts a file path, opens the workbook, and returns the name of the first worksheet. | Generate a C# snippet that loads a workbook, logs the number of worksheets, and gracefully handles missing‑file errors.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadDemo
{
    // Demonstrates how to instantiate a Workbook using the `Workbook(string)` constructor, check that the workbook contains at least one worksheet, output the worksheet count, and display the name of the first sheet for confirmation.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "sample.xlsx";

            // Load the workbook from the specified file path using the Workbook(string) constructor
            Workbook workbook = new Workbook(filePath);

            // Verify successful initialization by checking that the workbook contains at least one worksheet
            if (workbook.Worksheets != null && workbook.Worksheets.Count > 0)
            {
                Console.WriteLine($"Workbook loaded successfully. Worksheet count: {workbook.Worksheets.Count}");
            }
            else
            {
                Console.WriteLine("Failed to load workbook or workbook contains no worksheets.");
            }

            // Display the name of the first worksheet as additional confirmation
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine($"First worksheet name: {firstSheet.Name}");
        }
    }
}
