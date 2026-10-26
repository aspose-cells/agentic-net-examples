// Title: How to batch insert the same WAV OLE object into every worksheet of an Excel file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells add same OLE audio object to all worksheets | batch embed WAV file as OLE object in each sheet with Aspose.Cells | iterate through workbook worksheets to insert OLE object using Aspose.Cells .NET | Aspose.Cells OleObjects.Add multiple sheets example C# | programmatically add audio OLE object to every Excel worksheet in .NET
// Tags: batch insert OLE objects across worksheets Aspose.Cells | WAV audio OLE embedding Excel C# | OleObjects.Add method byte array Aspose.Cells | worksheet iteration for OLE insertion Aspose.Cells | programmatic audio OLE object Excel .NET

using System;
using System.IO;
using Aspose.Cells;

namespace OleObjectBatchProcess
{
    // The example loads an existing workbook, reads a WAV file into a byte array, loops through each worksheet, adds the WAV OLE object at a specified position using OleObjects.Add, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths for input workbook, OLE source file, and output workbook
                string inputFile = "input.xlsx";
                string wavFile = "sample.wav";
                string outputFile = "output.xlsx";

                // Verify that required files exist
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input workbook not found: {inputFile}");
                    return;
                }

                if (!File.Exists(wavFile))
                {
                    Console.WriteLine($"WAV file not found: {wavFile}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Read the WAV file into a byte array (required by OleObjects.Add)
                byte[] oleData = File.ReadAllBytes(wavFile);

                // Add the WAV OLE object to each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Parameters: upper left row, upper left column, height (pixels), width (pixels), OLE data
                    sheet.OleObjects.Add(0, 0, 100, 200, oleData);
                }

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
