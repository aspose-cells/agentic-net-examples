// Title: Save an Aspose.Cells workbook to CSV while preserving column separators for blank rows using TxtSaveOptions (C#)
// AI Prompts: Write C# code that creates a Workbook, adds data with intentional empty rows, and saves it as a CSV using TxtSaveOptions with KeepSeparatorsForBlankRow enabled. | Show how to configure TxtSaveOptions for ASCII encoding, a comma delimiter, and blank‑row separator preservation when exporting to CSV with Aspose.Cells. | Demonstrate exporting an Aspose.Cells workbook to CSV while ensuring separator characters are written for rows that contain no data.
// Common Searches: how to keep commas for empty rows when saving Excel to CSV with Aspose.Cells C# | Aspose.Cells TxtSaveOptions KeepSeparatorsForBlankRow C# example | save workbook as CSV preserving blank line delimiters Aspose.Cells | C# Aspose.Cells export CSV with custom separator and blank row handling
// Tags: Aspose.Cells CSV export KeepSeparatorsForBlankRow | TxtSaveOptions custom CSV settings C# | preserve blank row separators Aspose.Cells | ASCII encoding CSV Aspose.Cells | export workbook to CSV blank rows

using System;
using System.Text;
using Aspose.Cells;

// The example creates a workbook, inserts values with intentional empty rows, configures TxtSaveOptions to use ASCII encoding, a comma delimiter, and KeepSeparatorsForBlankRow=true, then saves the workbook as output.csv.
class SaveCsvWithBlankRowSeparators
{
    static void Main()
    {
        // Create a new workbook and get its cells collection
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Add data with intentional blank rows (rows 1 and 2 are left empty)
        cells[0, 0].PutValue("a");
        cells[0, 1].PutValue("b");
        cells[3, 0].PutValue("c");
        cells[4, 1].PutValue("d");

        // Configure TxtSaveOptions to keep separators for blank rows
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            Encoding = Encoding.ASCII,
            Separator = ',',               // Use comma as CSV delimiter
            KeepSeparatorsForBlankRow = true
        };

        // Save the workbook as CSV using the configured options
        workbook.Save("output.csv", saveOptions);
    }
}
