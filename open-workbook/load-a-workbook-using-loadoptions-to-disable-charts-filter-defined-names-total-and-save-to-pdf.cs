// Title: Load an Excel workbook, retain only the "Total" named range, and export it to PDF using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells remove all named ranges except a specific one before PDF export | How to keep only the 'Total' defined name when converting Excel to PDF with Aspose.Cells | Aspose.Cells filter NameCollection prior to saving workbook as PDF in .NET | Exclude unwanted named ranges from an Excel file using Aspose.Cells C# | Save Excel workbook as PDF after deleting selected defined names with Aspose.Cells
// Tags: filter defined names Aspose.Cells | named range cleanup before PDF export | Aspose.Cells workbook to PDF conversion | C# delete specific Excel named ranges | Excel defined name handling Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file, iterates through its defined names collection to remove every name except "Total", and then saves the modified workbook as a PDF, with error handling for missing files and load/save operations.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.pdf";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                Workbook workbook;
                try
                {
                    // Load the workbook (default load options are sufficient)
                    workbook = new Workbook(inputPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                    return;
                }

                // Access the collection of defined names
                NameCollection definedNames = workbook.Worksheets.Names;

                // Collect names to remove (all except "Total")
                var namesToRemove = new List<string>();
                foreach (Name name in definedNames)
                {
                    if (!string.Equals(name.Text, "Total", StringComparison.OrdinalIgnoreCase))
                    {
                        namesToRemove.Add(name.Text);
                    }
                }

                // Remove the unwanted names
                foreach (string name in namesToRemove)
                {
                    try
                    {
                        definedNames.Remove(name);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to remove name '{name}': {ex.Message}");
                    }
                }

                // Save the workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
