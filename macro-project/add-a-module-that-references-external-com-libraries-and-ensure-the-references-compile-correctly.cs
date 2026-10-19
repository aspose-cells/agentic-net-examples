// Title: Adding external COM library references to a C# Aspose.Cells project for reading and writing Excel cells
// Common Searches: how to add a COM reference to a .NET project that uses Aspose.Cells | C# compile error when referencing external COM libraries with Aspose.Cells | using Aspose.Cells to read cell A1 after adding COM interop in Visual Studio | step‑by‑step guide to include COM libraries in an Aspose.Cells console application | configure COM interop for Excel automation alongside Aspose.Cells in C#
// Tags: Aspose.Cells COM interop reference C# | read cell A1 Aspose.Cells .NET | create workbook save Xlsx Aspose.Cells | external COM library compilation Aspose.Cells | ExcelHelper utility class Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

namespace MyAsposeApp
{
    // Module that demonstrates reading and writing Excel files using Aspose.Cells.
    // This example shows how to add external COM library references to a C# project that uses Aspose.Cells, ensuring the project compiles correctly. The ExcelHelper class demonstrates creating a workbook, writing "Hello from Aspose" to cell A1, saving it as an XLSX file, and reading the value back, with proper error handling and console output.
    public static class ExcelHelper
    {
        // Reads the value of cell A1 from an existing Excel file using Aspose.Cells.
        public static string ReadCellA1(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.");

            try
            {
                // Load the workbook.
                var workbook = new Workbook(filePath);
                // Access the first worksheet.
                var worksheet = workbook.Worksheets[0];
                // Get the value of cell A1.
                var value = worksheet.Cells["A1"].Value?.ToString() ?? string.Empty;
                return value;
            }
            catch (Exception ex)
            {
                // Log or rethrow as needed.
                Console.Error.WriteLine($"Error reading cell A1: {ex.Message}");
                throw;
            }
        }

        // Example method that creates a workbook with Aspose.Cells and then reads it.
        public static void CreateAndRead(string filePath)
        {
            try
            {
                // Ensure the directory exists.
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                // Create a new workbook and set a value in A1.
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Hello from Aspose");
                workbook.Save(filePath, SaveFormat.Xlsx);

                // Read the same file using Aspose.Cells.
                string readValue = ReadCellA1(filePath);
                Console.WriteLine($"Value read: {readValue}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error in CreateAndRead: {ex.Message}");
                throw;
            }
        }
    }

    // Entry point for the application.
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                // Determine file path: use first argument or default to Sample.xlsx in current directory.
                string filePath = args.Length > 0
                    ? args[0]
                    : Path.Combine(Directory.GetCurrentDirectory(), "Sample.xlsx");

                if (string.IsNullOrWhiteSpace(filePath))
                {
                    Console.Error.WriteLine("Invalid file path.");
                    return;
                }

                // Execute the example.
                ExcelHelper.CreateAndRead(filePath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
