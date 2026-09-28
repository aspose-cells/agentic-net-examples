// Title: How to explicitly dispose Aspose.Cells Workbook objects in C# to release file handles and free memory
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, updates a specific cell, saves the workbook, and guarantees Workbook.Dispose is called in a finally block. | Create a reusable C# method for processing multiple workbooks that writes values, saves each file, and ensures each Workbook instance is disposed even when exceptions occur.
// Common Searches: C# Aspose.Cells how to release file lock after saving workbook | ensure Workbook.Dispose is called when using Aspose.Cells in a loop | best practice for disposing Aspose.Cells Workbook objects in .NET | avoid Excel file handle leak with Aspose.Cells workbook | using try‑finally to call Workbook.Dispose in Aspose.Cells C#
// Tags: Aspose.Cells workbook disposal C# | release Excel file handles Aspose.Cells | workbook lifecycle management Aspose.Cells | explicit Workbook.Dispose usage | memory cleanup after saving Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

// // This program processes two Excel files with Aspose.Cells, writes values to specified cells, saves each workbook, and explicitly calls Workbook.Dispose to release file handles and free memory.
class Program
{
    static void Main()
    {
        // Process first workbook
        ProcessWorkbook("Input1.xlsx", "Output1.xlsx", "A1", "Hello World");
        // Process second workbook
        ProcessWorkbook("Input2.xlsx", "Output2.xlsx", "B2", 12345);
    }

    static void ProcessWorkbook(string inputPath, string outputPath, string cellName, object value)
    {
        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Write the value to the specified cell in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells[cellName].PutValue(value);

            // Save the workbook to the output file
            workbook.Save(outputPath);
            workbook.Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing '{inputPath}': {ex.Message}");
        }
    }
}
