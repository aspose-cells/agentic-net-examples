// Title: Load an existing Excel workbook from a file path using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells load workbook from local file path | How to open an existing .xlsx file with Aspose.Cells for .NET | Reading Excel file into Aspose.Cells Workbook object in C# | Validate Excel file existence before loading with Aspose.Cells | Get worksheet names after loading Excel workbook using Aspose.Cells C# | Determine workbook file format after opening with Aspose.Cells
// Tags: Aspose.Cells workbook initialization from file system | C# verify file path before Aspose.Cells workbook creation | Aspose.Cells retrieve worksheet collection after load | Aspose.Cells query workbook file format property | load Excel workbook from disk using Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The sample program checks that a command‑line argument points to an existing file, then creates an Aspose.Cells Workbook by passing the file path to its constructor. It prints the workbook's file format, the total number of worksheets, and each worksheet's name before disposing the workbook.
class Program
{
    static void Main(string[] args)
    {
        // Verify that a file path was provided
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: Program <excel-file-path>");
            return;
        }

        string filePath = args[0];

        // Check that the file exists
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" does not exist.");
            return;
        }

        try
        {
            // Load the Excel file into a Workbook instance
            Workbook workbook = new Workbook(filePath);

            // Output basic information about the loaded workbook
            Console.WriteLine("Workbook loaded successfully.");
            Console.WriteLine($"File format: {workbook.FileFormat}");
            Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");

            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Console.WriteLine($"  Worksheet {i + 1}: {workbook.Worksheets[i].Name}");
            }

            // Dispose the workbook when done (optional, as using statement could be used)
            workbook.Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
        }
    }
}
