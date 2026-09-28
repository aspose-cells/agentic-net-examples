// Title: Rename Excel named ranges that start with "Temp" by adding an "Archive_" prefix using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, iterates over workbook.Worksheets.Names, and prepends "Archive_" to any Name.Text that begins with "Temp", then saves the file. | Generate a loop that examines each defined name in a workbook and updates its Text property to add an "Archive_" prefix when the name starts with "Temp" using the Aspose.Cells API. | Create a console application that accepts input and output Excel paths, renames temporary named ranges by prefixing them with "Archive_" via Aspose.Cells, and includes error handling for missing files.
// Common Searches: Aspose.Cells C# rename defined names that start with Temp | Add prefix to Excel named ranges using Aspose.Cells .NET | Loop through workbook.Worksheets.Names and change Name.Text in C# | How to archive temporary named ranges in an Excel file with Aspose.Cells | Programmatically rename multiple named ranges in a .NET Excel library
// Tags: rename named ranges Aspose.Cells | prefix defined names C# | modify Name.Text property Aspose.Cells | batch update Excel named ranges .NET | archive temporary ranges Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, iterates through all defined names, and prefixes any name starting with "Temp" with "Archive_" using Aspose.Cells for .NET, then saves the updated workbook.
class Program
{
    static void Main(string[] args)
    {
        // Determine input and output file paths (use command‑line arguments or defaults)
        string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
        string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified input file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all defined names in the workbook
            int nameCount = workbook.Worksheets.Names.Count;
            for (int i = 0; i < nameCount; i++)
            {
                Name definedName = workbook.Worksheets.Names[i];

                // If the name starts with "Temp", prefix it with "Archive_"
                if (definedName.Text.StartsWith("Temp", StringComparison.Ordinal))
                {
                    definedName.Text = "Archive_" + definedName.Text;
                }
            }

            // Save the modified workbook to the specified output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
