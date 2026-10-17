// Title: C# custom iterator to enumerate only visible rows in an Aspose.Cells worksheet
// AI Prompts: Generate a C# GetVisibleRows method that yields each Row from a Worksheet where Row.IsHidden is false using Aspose.Cells. | Write code that creates a VisibleRowEnumerator class for an Aspose.Cells worksheet and prints the index of every non‑hidden row. | Show how to implement an iterator that skips hidden rows while looping through Worksheet.Cells.Rows in a .NET Excel processing routine.
// Common Searches: how to iterate over only visible rows in an Aspose.Cells worksheet using C# | C# example for filtering out hidden rows with Aspose.Cells | enumerate non‑hidden rows in Excel file with Aspose.Cells for .NET | custom iterator to skip hidden rows in Aspose.Cells workbook | retrieve visible row indices from a worksheet using Aspose.Cells API
// Tags: enumerate visible rows Aspose.Cells | skip hidden rows iterator C# | filter worksheet rows by IsHidden | custom Excel row enumerator .NET | process non‑hidden rows Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example defines a VisibleRowEnumerator class with a GetVisibleRows method that yields only rows whose IsHidden property is false from a given Worksheet. It loads an Excel workbook, creates the enumerator, and writes each visible row's index to the console.
public class VisibleRowEnumerator
{
    private readonly Worksheet _worksheet;

    public VisibleRowEnumerator(Worksheet worksheet)
    {
        _worksheet = worksheet;
    }

    // Iterator that yields rows which are not hidden
    public IEnumerable<Row> GetVisibleRows()
    {
        // Iterate through all rows in the worksheet
        foreach (Row row in _worksheet.Cells.Rows)
        {
            // Skip rows that are hidden
            if (row.IsHidden)
                continue;

            yield return row;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Path to the input workbook
        string inputPath = "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Create the visible row enumerator
            var enumerator = new VisibleRowEnumerator(sheet);

            // Process each visible row
            foreach (Row visibleRow in enumerator.GetVisibleRows())
            {
                // Example processing: output the row index
                Console.WriteLine($"Visible Row Index: {visibleRow.Index}");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
