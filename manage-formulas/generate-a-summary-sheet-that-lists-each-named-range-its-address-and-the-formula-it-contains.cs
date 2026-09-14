// Title: Create a summary worksheet that lists each named range, its address, and its formula using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an existing workbook, adds a worksheet named "Summary", and fills columns A‑C with the name, RefersTo address, and formula (if any) of every defined name. | Implement a method that loops through workbook.Worksheets.Names, extracts definedName.Text, definedName.RefersTo, detects formulas starting with '=', writes the data to a new sheet, auto‑fits the columns, and saves the file.
// Common Searches: asp.net aspose.cells list all named ranges with their formulas in a new worksheet | c# example to export defined names and RefersTo addresses to a summary sheet using Aspose.Cells | how to generate a report of named ranges including address and formula with Aspose.Cells .NET | sample code for iterating workbook.Worksheets.Names and writing results to a summary worksheet
// Tags: Aspose.Cells list named ranges | Aspose.Cells export defined names to worksheet | Aspose.Cells write RefersTo address | Aspose.Cells create summary sheet | Aspose.Cells auto fit columns

using Aspose.Cells;
using System;
using System.IO;

// The program loads an existing Excel file, adds a "Summary" worksheet, writes each defined name's text, its RefersTo address, and any leading‑"=" formula into columns A‑C, auto‑fits the columns, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook that contains named ranges.
            Workbook workbook = new Workbook(inputPath);

            // Add a new worksheet for the summary and name it "Summary".
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Write header titles.
            summarySheet.Cells["A1"].PutValue("Named Range");
            summarySheet.Cells["B1"].PutValue("Address");
            summarySheet.Cells["C1"].PutValue("Formula");

            int currentRow = 1; // Zero‑based index; row 2 in the sheet.

            // Iterate through all defined names (named ranges) in the workbook.
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                // Column A: the name of the defined range.
                summarySheet.Cells[currentRow, 0].PutValue(definedName.Text);

                // Column B: the address the name refers to (e.g., Sheet1!$A$1:$B$5).
                // For formula names, this will contain the formula string.
                summarySheet.Cells[currentRow, 1].PutValue(definedName.RefersTo);

                // Column C: if the name is a formula, display the formula; otherwise leave blank.
                // Aspose.Cells older versions may not expose IsFormula, so infer from the RefersTo string.
                string formula = (!string.IsNullOrEmpty(definedName.RefersTo) && definedName.RefersTo.StartsWith("="))
                                 ? definedName.RefersTo
                                 : string.Empty;
                summarySheet.Cells[currentRow, 2].PutValue(formula);

                currentRow++;
            }

            // Adjust column widths to fit the content.
            summarySheet.AutoFitColumns();

            // Save the workbook with the new summary sheet.
            workbook.Save(outputPath);
            Console.WriteLine($"Summary sheet created and saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
