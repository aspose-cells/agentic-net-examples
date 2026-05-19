using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class OleObjectIsLinkDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths (adjust as needed)
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                Run(inputPath, outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Core processing method
        public static void Run(string inputFile, string outputFile)
        {
            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all OLE objects in the current worksheet
                    foreach (OleObject ole in sheet.OleObjects)
                    {
                        // Identify linked OLE objects using the IsLink property
                        if (ole.IsLink)
                        {
                            // Display source file path of linked OLE object
                            Console.WriteLine($"Linked OLE object found in sheet \"{sheet.Name}\": {ole.ObjectSourceFullName}");
                            
                            // Example modification (optional):
                            // string newPath = ole.ObjectSourceFullName.Replace(@"C:\", @"D:\");
                            // ole.ObjectSourceFullName = newPath;
                        }
                        else
                        {
                            // Skip embedded OLE objects
                            Console.WriteLine($"Embedded OLE object skipped in sheet \"{sheet.Name}\".");
                        }
                    }
                }

                // Save the workbook after processing
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved as {outputFile}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}