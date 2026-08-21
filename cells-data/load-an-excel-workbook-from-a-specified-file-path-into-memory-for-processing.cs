// Title: Load an Excel workbook from a file path using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open an Excel file with Aspose.Cells by passing a full file path to the Workbook constructor, validate the file's existence, handle loading errors, and retrieve basic worksheet information such as the first sheet name and total sheet count.
// Keywords: Aspose.Cells load workbook | C# open Excel file | Workbook constructor file path | validate Excel file existence | exception handling Aspose.Cells | read worksheet names C# | Aspose.Cells .NET example
// Common Searches: How to open an Excel file with Aspose.Cells in C# | Aspose.Cells load workbook from path | C# read first worksheet name using Aspose.Cells | Count worksheets after loading workbook Aspose | Aspose.Cells error handling when file not found
// Developer Intent: Load an Excel workbook from a specified path into memory so it can be processed programmatically.
// Use Cases: Open a user‑provided Excel file for data extraction or transformation. | Verify that the target file exists before creating a Workbook object. | Capture and log exceptions that occur during workbook initialization. | Retrieve the name of the first worksheet and the total number of sheets for quick validation.
// AI Prompts: Write C# code that uses Aspose.Cells to load an Excel file from a given path, prints all worksheet names, and gracefully handles missing‑file or corrupted‑file errors. | Provide a reusable method in C# for opening an Excel workbook with Aspose.Cells that includes file‑existence checks and detailed exception logging.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to open an Excel file with Aspose.Cells by passing a full file path to the Workbook constructor, validate the file's existence, handle loading errors, and retrieve basic worksheet information such as the first sheet name and total sheet count.
    public class LoadWorkbookDemo
    {
        /// <param name="filePath">Full path to the Excel file to load.</param>
        public static void Run(string filePath)
        {
            try
            {
                // Load the workbook using the constructor that accepts a file path.
                Workbook workbook = new Workbook(filePath);

                // Example processing: display the name of the first worksheet and total sheet count.
                Worksheet firstSheet = workbook.Worksheets[0];
                Console.WriteLine($"First worksheet name: {firstSheet.Name}");
                Console.WriteLine($"Total worksheets: {workbook.Worksheets.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            string filePath;

            if (args.Length > 0)
            {
                filePath = args[0];
            }
            else
            {
                // Provide a default path or prompt the user.
                Console.Write("Enter the full path to the Excel file: ");
                filePath = Console.ReadLine();
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                Console.WriteLine("No file path provided.");
                return;
            }

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Run the demo with the validated file path.
            LoadWorkbookDemo.Run(filePath);
        }
    }
}
