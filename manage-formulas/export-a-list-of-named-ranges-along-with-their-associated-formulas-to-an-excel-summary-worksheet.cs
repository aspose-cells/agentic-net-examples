// Title: Export named ranges and their RefersTo formulas to a Summary worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds a worksheet named 'Summary', writes each workbook named range and its RefersTo formula into columns A and B, auto‑fits the columns, and saves the file. | Generate a method that iterates Workbook.Worksheets.Names, records the Name.Text and Name.RefersTo values on a new sheet, and returns the updated Workbook. | Create a script that checks for an existing 'Summary' sheet, clears it if present, populates it with named range details, and writes the workbook to a specified output path.
// Common Searches: Aspose.Cells C# how to list all named ranges with their formulas in a new sheet | export named range RefersTo property to a summary worksheet using Aspose.Cells for .NET | C# code to create a summary tab that shows workbook named ranges and formulas | iterate workbook.Worksheets.Names and write results to another worksheet Aspose.Cells | auto fit columns after writing named range data Aspose.Cells C#
// Tags: export named ranges to summary worksheet Aspose.Cells | retrieve RefersTo formula from named range C# | create summary sheet with named range details Aspose.Cells | list workbook named ranges and formulas .NET | auto‑fit columns after writing data Aspose.Cells | clear and reuse existing summary worksheet C#

using Aspose.Cells;
using System;
using System.IO;

// Loads an existing workbook (or creates a new one), ensures a 'Summary' worksheet, writes headers, iterates all defined names, records each name and its RefersTo formula into columns A and B, auto‑fits columns, and saves the workbook to the output file.
class ExportNamedRanges
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists; otherwise create an empty workbook
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Ensure a summary worksheet exists (create if not)
            Worksheet summarySheet = workbook.Worksheets["Summary"];
            if (summarySheet != null)
            {
                // Clear any previous content
                summarySheet.Cells.Clear();
            }
            else
            {
                summarySheet = workbook.Worksheets.Add("Summary");
            }

            // Write header titles
            summarySheet.Cells["A1"].PutValue("Named Range");
            summarySheet.Cells["B1"].PutValue("Refers To (Formula)");

            // Export each named range and its formula
            int row = 1; // zero‑based index (row 2 in Excel)
            foreach (Name namedRange in workbook.Worksheets.Names)
            {
                // Column A: name of the range
                summarySheet.Cells[row, 0].PutValue(namedRange.Text);
                // Column B: the formula the name refers to
                summarySheet.Cells[row, 1].PutValue(namedRange.RefersTo);
                row++;
            }

            // Adjust column widths for readability
            summarySheet.AutoFitColumns();

            // Save the workbook with the new summary sheet
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
