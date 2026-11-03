// Title: Create a C# custom iterator that enumerates only visible rows in an Aspose.Cells worksheet
// Common Searches: C# how to skip hidden rows when looping through Aspose.Cells worksheet rows | custom IEnumerable to filter visible rows in an Excel file using Aspose.Cells | enumerate only non‑hidden rows with Aspose.Cells API in C# | iterate over visible rows in a workbook worksheet without processing hidden rows | Aspose.Cells C# example for Row.IsHidden based row filtering
// Tags: visible rows iterator Aspose.Cells | skip hidden rows C# enumerable | filter worksheet rows by Row.IsHidden | custom IEnumerable for Excel rows | process only visible rows in workbook

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHiddenRowIterator
{
    // Custom enumerable that iterates only over visible rows in a worksheet
    // The sample defines a VisibleRowEnumerable that implements IEnumerable<Row> and yields rows from a worksheet only when Row.IsHidden is false. It demonstrates loading a workbook, obtaining a worksheet, and using the custom iterator to process each visible row (e.g., printing its index) before optionally saving the file.
    public class VisibleRowEnumerable : IEnumerable<Row>
    {
        private readonly Worksheet _worksheet;

        public VisibleRowEnumerable(Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public IEnumerator<Row> GetEnumerator()
        {
            // Iterate through all rows in the worksheet's Cells collection
            foreach (Row row in _worksheet.Cells.Rows)
            {
                // Skip rows that are hidden
                if (!row.IsHidden)
                {
                    yield return row;
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook (no password supplied). If the file is password‑protected,
                // a CellsException will be thrown and handled in the catch block.
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or any specific worksheet)
                Worksheet worksheet = workbook.Worksheets[0];

                // Use the custom iterator to process only visible rows
                foreach (Row visibleRow in new VisibleRowEnumerable(worksheet))
                {
                    // Example processing: print the row index
                    Console.WriteLine($"Processing visible row: {visibleRow.Index}");
                    // Additional logic can be placed here
                }

                // Save the workbook if any modifications were made (optional)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (CellsException ex) when (ex.Message.Contains("password"))
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a valid password to open it.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
