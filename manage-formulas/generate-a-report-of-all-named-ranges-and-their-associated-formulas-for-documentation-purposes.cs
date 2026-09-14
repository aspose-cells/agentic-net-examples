// Title: C# Aspose.Cells example to generate an Excel report of all named ranges with their RefersTo formulas
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, loops through every defined name, and writes the name and its RefersTo formula into a new worksheet. | Enhance the program to also capture whether each named range is scoped to a specific worksheet or the whole workbook and add a column for the scope in the report. | Convert the output format to CSV, producing a file that lists the named range, its RefersTo formula, and its scope.
// Common Searches: how to list all named ranges and their formulas with Aspose.Cells in C# | Aspose.Cells C# export defined names to a separate Excel file | C# generate documentation of Excel named ranges using Aspose.Cells | retrieve RefersTo property of named ranges Aspose.Cells example | create a named range report workbook programmatically with Aspose.Cells
// Tags: Aspose.Cells enumerate defined names | export named ranges to Excel report C# | retrieve RefersTo formula Aspose.Cells | named range scope detection Aspose.Cells | generate CSV from named ranges Aspose.Cells

using System;
using Aspose.Cells;

namespace NamedRangeReport
{
    // // Loads a source workbook, iterates through its defined names, records each name, its RefersTo formula (and optionally its scope) into a new worksheet, auto‑fits columns, and saves the result as a separate report file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (adjust as needed)
            string inputFilePath = "input.xlsx";
            string outputFilePath = "NamedRangesReport.xlsx";

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputFilePath);

            // Create a new workbook for the report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];

            // Write header row
            reportSheet.Cells["A1"].PutValue("Named Range");
            reportSheet.Cells["B1"].PutValue("Refers To (Formula)");

            // Start writing data from the second row
            int currentRow = 1; // zero‑based index (row 2 in Excel)

            // Iterate through all defined names (named ranges) in the source workbook
            foreach (Name namedRange in sourceWorkbook.Worksheets.Names)
            {
                // Get the name and the formula it refers to
                string name = namedRange.Text;          // e.g., "MyRange"
                string formula = namedRange.RefersTo;   // e.g., "=Sheet1!$A$1:$A$10"

                // Write the data into the report sheet
                reportSheet.Cells[currentRow, 0].PutValue(name);
                reportSheet.Cells[currentRow, 1].PutValue(formula);

                currentRow++;
            }

            // Auto‑fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the report workbook
            reportWorkbook.Save(outputFilePath);
        }
    }
}
