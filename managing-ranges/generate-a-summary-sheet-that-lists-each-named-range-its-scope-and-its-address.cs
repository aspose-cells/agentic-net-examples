// Title: Create a summary worksheet that lists every named range with its scope and address using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that adds a new worksheet called 'Summary', writes headers, and fills rows with each named range’s name, its default workbook scope, and the RefersTo address. | Develop a program that loads an existing Excel file, iterates over workbook.Worksheets.Names, and exports the name, scope, and address of each named range to columns A‑C of a newly created sheet, then saves the workbook.
// Common Searches: Aspose.Cells C# list all named ranges and their RefersTo addresses | How to generate a summary sheet of named ranges with scope using Aspose.Cells .NET | Retrieve named range scope and address programmatically with Aspose.Cells | Export named range details to a new worksheet in C# using Aspose.Cells | Create a summary tab that shows named range names, scopes, and addresses in Excel via Aspose.Cells
// Tags: Aspose.Cells enumerate named ranges | export named range details to worksheet | list named range address .NET | generate summary sheet for named ranges | retrieve named range scope Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, adds a worksheet named 'Summary', writes column headers, loops through workbook.Worksheets.Names, and writes each named range’s name, a default 'Workbook' scope, and its RefersTo address into columns A‑C before saving the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Add a new worksheet for the summary and name it "Summary"
            int summaryIndex = workbook.Worksheets.Add();
            var summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Write header titles
            summarySheet.Cells["A1"].PutValue("Named Range");
            summarySheet.Cells["B1"].PutValue("Scope");
            summarySheet.Cells["C1"].PutValue("Address");

            int row = 1; // Start from the second row (zero‑based index)

            // Iterate through all named ranges in the workbook
            foreach (Name namedRange in workbook.Worksheets.Names)
            {
                // Column A: Named range name
                summarySheet.Cells[row, 0].PutValue(namedRange.Text);

                // Column B: Scope – default to "Workbook" (global). 
                // Aspose.Cells older versions may not expose scope details.
                string scope = "Workbook";
                summarySheet.Cells[row, 1].PutValue(scope);

                // Column C: Address the name refers to
                summarySheet.Cells[row, 2].PutValue(namedRange.RefersTo);

                row++;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
