// Title: How to add external COM library references to a C# Aspose.Cells project and generate an Excel workbook
// Common Searches: add COM reference to a .NET project that uses Aspose.Cells | compile Aspose.Cells C# code with external COM interop libraries | generate Excel file with Aspose.Cells after configuring COM references | C# Aspose.Cells example that requires external COM component reference | resolve missing COM library errors in an Aspose.Cells macro project
// Tags: Aspose.Cells COM reference configuration | C# generate Excel workbook Aspose.Cells | external COM interop setup .NET | Workbook.Save custom file path C# | exception handling Aspose.Cells Excel creation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsComIntegration
{
    // This module demonstrates creating an Excel file using Aspose.Cells.
    // The example shows how to structure a C# module that creates an Excel workbook with Aspose.Cells, ensures the output directory exists, writes a value to cell A1, and saves the file. It also outlines the steps needed to reference external COM libraries in the project so that the code compiles without errors, and includes basic exception handling for the file creation process.
    public static class ExternalComModule
    {
        // Creates an Excel file, writes a value to cell A1, and saves it.
        public static void CreateExcelFile(string filePath)
        {
            try
            {
                // Ensure the output directory exists.
                string? directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Create a new workbook.
                var workbook = new Workbook();

                // Access the first worksheet.
                var worksheet = workbook.Worksheets[0];

                // Write a value to cell A1.
                worksheet.Cells["A1"].PutValue("Hello from Aspose.Cells!");

                // Save the workbook to the specified path.
                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                // Log the exception and rethrow for the caller to handle.
                Console.Error.WriteLine($"Error creating Excel file: {ex.Message}");
                throw;
            }
        }
    }

    // Simple entry point for demonstration purposes.
    public class Program
    {
        public static void Main(string[] args)
        {
            string outputPath = "Output.xlsx";

            try
            {
                ExternalComModule.CreateExcelFile(outputPath);
                Console.WriteLine($"Excel file successfully created at: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to create Excel file: {ex.Message}");
            }
        }
    }
}
