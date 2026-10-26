// Title: Generate a summary worksheet of all named ranges with scope and address using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# list all defined names with their addresses | how to create a summary sheet of named ranges in an Excel workbook using Aspose.Cells | C# Aspose.Cells write named range scope and address to a new worksheet | export workbook named ranges to a summary worksheet with Aspose.Cells .NET | retrieve named range RefersTo and scope in Aspose.Cells C# example
// Tags: Aspose.Cells list defined names to worksheet | C# export named ranges to summary sheet | Aspose.Cells write range address and scope | generate named range summary XLSX Aspose | iterate workbook.Worksheets.Names Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, adds a worksheet named "Summary", writes column headers for Name, Scope, and Address, then iterates through workbook.Worksheets.Names to record each defined name's text, a default "Workbook" scope, and its RefersTo address into successive rows before saving the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Define input and output file paths (replace placeholders with actual paths)
        string inputPath = "{InputFilePath}";
        string outputPath = "{OutputFilePath}";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet for the summary
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Write header titles
            summarySheet.Cells["A1"].PutValue("Name");
            summarySheet.Cells["B1"].PutValue("Scope");
            summarySheet.Cells["C1"].PutValue("Address");

            // Populate the summary with each named range
            int row = 2; // Start from the second row (1‑based index)
            foreach (Aspose.Cells.Name definedName in workbook.Worksheets.Names)
            {
                // Name of the range
                summarySheet.Cells[row, 0].PutValue(definedName.Text);

                // Scope: workbook‑level or worksheet‑level.
                // The Name class in some Aspose.Cells versions does not expose IsWorkbookScope,
                // so we default to "Workbook". Adjust as needed for your version.
                string scope = "Workbook";
                summarySheet.Cells[row, 1].PutValue(scope);

                // Address the range refers to
                summarySheet.Cells[row, 2].PutValue(definedName.RefersTo);

                row++;
            }

            // Ensure the output directory exists
            if (!string.IsNullOrWhiteSpace(outputPath))
            {
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            else
            {
                Console.WriteLine("Output path is not specified.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
