// Title: How to retain only the "Summary" and "Data" worksheets when processing an Excel file with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells keep only selected worksheets after loading workbook | remove all sheets except Summary and Data using Aspose.Cells .NET | filter worksheets in an Excel file with Aspose.Cells before saving | Aspose.Cells C# example to delete unwanted worksheets from workbook | how to load Excel file and discard specific sheets using Aspose.Cells
// Tags: Aspose.Cells selective worksheet removal C# | filter worksheets after workbook load Aspose.Cells | retain specific Excel sheets Aspose.Cells .NET | remove unwanted worksheets Aspose.Cells C# | C# workbook worksheet filtering Aspose.Cells | Excel sheet exclusion using Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program verifies the source Excel file exists, loads it into a Workbook, defines an array with the allowed sheet names ("Summary" and "Data"), iterates the worksheets in reverse order removing any that are not in the allowed list, prints the names of the retained worksheets, and saves the filtered workbook to a new file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel file
                string sourceFile = "input.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file not found: {sourceFile}");
                    return;
                }

                // Load the workbook normally
                Workbook workbook = new Workbook(sourceFile);

                // Names of worksheets that should be retained
                string[] allowedSheets = { "Summary", "Data" };

                // Remove worksheets that are not in the allowed list
                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    bool keep = false;
                    foreach (string name in allowedSheets)
                    {
                        if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                        {
                            keep = true;
                            break;
                        }
                    }

                    if (!keep)
                    {
                        workbook.Worksheets.RemoveAt(i);
                    }
                }

                // Display the names of the retained worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Console.WriteLine($"Loaded worksheet: {sheet.Name}");
                }

                // Save the filtered workbook to a new file
                string outputFile = "filtered_output.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Filtered workbook saved to: {outputFile}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
