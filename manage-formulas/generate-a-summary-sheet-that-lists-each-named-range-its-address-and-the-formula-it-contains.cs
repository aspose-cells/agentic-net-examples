// Title: Generate a summary worksheet of all named ranges with their addresses and formulas using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create summary worksheet for named ranges with address and formula | How to export named range definitions to a new sheet using Aspose.Cells .NET | C# code to list workbook named ranges and their RefersTo values with Aspose.Cells | Generate Excel report of named ranges including formulas using Aspose.Cells for .NET | Extract named range addresses and formulas and write to a summary tab in C#
// Tags: Aspose.Cells generate named range overview | C# retrieve named range RefersTo content | Aspose.Cells list workbook named ranges | C# add overview worksheet for named ranges | detect formula based named ranges Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook with Aspose.Cells, adds a new worksheet named "Summary", writes column headers for Named Range, Address, and Formula, then iterates through all workbook named ranges. For each name it records the name text and, depending on whether the RefersTo string starts with '=', writes either the address or the formula (leaving the other column blank). Finally, the workbook is saved to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet for the summary and name it "Summary"
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Write header titles
            summarySheet.Cells[0, 0].PutValue("Named Range");
            summarySheet.Cells[0, 1].PutValue("Address");
            summarySheet.Cells[0, 2].PutValue("Formula");

            // Start writing data from the second row
            int currentRow = 1;

            // Iterate through all named ranges in the workbook
            foreach (Name namedRange in workbook.Worksheets.Names)
            {
                // Column A: the name of the range
                summarySheet.Cells[currentRow, 0].PutValue(namedRange.Text);

                // Determine if the name refers to a formula (RefersTo starts with '=')
                bool isFormula = !string.IsNullOrEmpty(namedRange.RefersTo) && namedRange.RefersTo.StartsWith("=");

                if (!isFormula)
                {
                    // Non‑formula name: write the address, leave formula column empty
                    summarySheet.Cells[currentRow, 1].PutValue(namedRange.RefersTo);
                    summarySheet.Cells[currentRow, 2].PutValue(string.Empty);
                }
                else
                {
                    // Formula name: leave address empty and write the formula
                    summarySheet.Cells[currentRow, 1].PutValue(string.Empty);
                    summarySheet.Cells[currentRow, 2].PutValue(namedRange.RefersTo);
                }

                currentRow++;
            }

            // Save the modified workbook with the summary sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
