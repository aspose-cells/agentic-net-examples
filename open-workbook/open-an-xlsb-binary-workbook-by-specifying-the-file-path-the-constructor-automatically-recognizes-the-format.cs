// Title: Open an XLSB Workbook from a File Path with Aspose.Cells for .NET (C#)
// Description: Shows how to load a binary XLSB workbook in C# by supplying its file path to the Aspose.Cells Workbook constructor, which automatically detects the format. If the file does not exist, the sample creates a minimal XLSB workbook, reads the value of cell A1, and prints the total number of worksheets.
// Keywords: Aspose.Cells | XLSB | C# | .NET | load workbook from path | automatic format detection | read cell value | create sample XLSB | worksheet count | binary Excel file
// Common Searches: Aspose.Cells open XLSB C# | C# load binary Excel workbook with Aspose.Cells | Workbook constructor file path format detection | Create sample XLSB file if missing Aspose.Cells | Read cell A1 from XLSB using Aspose.Cells
// Developer Intent: Load an existing XLSB file (or generate a fallback file) and retrieve cell data or worksheet information using Aspose.Cells in a .NET project.
// Use Cases: Open a production XLSB report and extract specific cell values. | Automatically generate a placeholder XLSB file when the expected file is absent. | Count worksheets in a binary workbook to validate file structure before processing. | Integrate XLSB loading into a server‑side C# service that handles Excel data.
// AI Prompts: Generate C# code that opens an XLSB workbook from a given path with Aspose.Cells, creating a simple workbook if the file is missing. | Provide a snippet that reads cell A1 and prints the total worksheet count from an XLSB file using Aspose.Cells for .NET. | Explain how the Aspose.Cells Workbook constructor determines the file format without explicit parameters. | Troubleshoot errors that occur when loading a corrupted XLSB file with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXlsbOpenExample
{
    // Shows how to load a binary XLSB workbook in C# by supplying its file path to the Aspose.Cells Workbook constructor, which automatically detects the format. If the file does not exist, the sample creates a minimal XLSB workbook, reads the value of cell A1, and prints the total number of worksheets.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSB file.
            string xlsbPath = @"C:\Data\sample.xlsb";

            // Ensure the file exists; create a simple workbook if it does not.
            if (!File.Exists(xlsbPath))
            {
                Console.WriteLine($"File not found: {xlsbPath}");
                try
                {
                    Workbook newWb = new Workbook();
                    newWb.Worksheets[0].Cells["A1"].PutValue("Sample");
                    newWb.Save(xlsbPath, SaveFormat.Xlsb);
                    Console.WriteLine("A sample XLSB workbook has been created.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                    return;
                }
            }

            // Load the workbook safely.
            Workbook workbook;
            try
            {
                workbook = new Workbook(xlsbPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Read and display the value of cell A1.
            Cell cell = worksheet.Cells["A1"];
            Console.WriteLine($"Cell A1 Value: {cell.Value}");

            // Display the total number of worksheets.
            Console.WriteLine($"Total Worksheets: {workbook.Worksheets.Count}");
        }
    }
}
