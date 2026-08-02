// Title: Export Worksheet Error‑Check Settings to a Text Report with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook using Aspose.Cells for .NET, reads each worksheet's ErrorCheckOptionCollection, enumerates enabled ErrorCheckType values and their target ranges, and writes a plain‑text report that lists the worksheet name, option index, active checks, and cell areas. The workbook can be saved unchanged after reporting.
// Keywords: Aspose.Cells | C# error check options | ErrorCheckOptionCollection | read worksheet error checking | export error check settings | Excel error checking report | Aspose.Cells .NET | ErrorCheckType enumeration | log Excel error checks | generate text report
// Common Searches: Aspose.Cells read error check options C# | How to list enabled error checks in Excel using Aspose.Cells | Export worksheet error‑checking configuration to text file | C# code to iterate ErrorCheckOptionCollection | Generate report of Excel error checking rules with Aspose.Cells
// Developer Intent: Read each worksheet’s error‑checking configuration and produce a human‑readable text file that details the enabled checks and their applied ranges.
// Use Cases: Audit custom error‑checking rules before distributing a workbook | Document error‑check settings for compliance or review purposes | Compare error‑checking configurations across different workbook versions | Automate generation of documentation for data‑validation policies
// AI Prompts: Write a C# method that takes a Workbook object and returns the error‑check report as a string instead of writing to a file. | Show how to filter the report to include only specific ErrorCheckType values such as NumberAsText and InconsistentFormula. | Explain how to modify the code to export the report as CSV with columns: Worksheet, OptionIndex, EnabledCheck, RangeStart, RangeEnd.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook using Aspose.Cells for .NET, reads each worksheet's ErrorCheckOptionCollection, enumerates enabled ErrorCheckType values and their target ranges, and writes a plain‑text report that lists the worksheet name, option index, active checks, and cell areas. The workbook can be saved unchanged after reporting.
class ErrorCheckReport
{
    static void Main()
    {
        // Load an existing workbook (load rule)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Path for the text report
        string reportPath = "ErrorCheckReport.txt";

        // Create a writer for the report file
        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            // Iterate through all worksheets in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                writer.WriteLine($"Worksheet: {ws.Name}");

                // Access the collection of error‑check options for the worksheet
                ErrorCheckOptionCollection options = ws.ErrorCheckOptions;

                // If there are no options, note it and continue
                if (options.Count == 0)
                {
                    writer.WriteLine("  No error‑check options defined.");
                    writer.WriteLine();
                    continue;
                }

                // Process each ErrorCheckOption in the collection
                for (int i = 0; i < options.Count; i++)
                {
                    ErrorCheckOption option = options[i];
                    writer.WriteLine($"  Option #{i + 1}:");

                    // List all enabled error‑check types for this option
                    foreach (ErrorCheckType type in Enum.GetValues(typeof(ErrorCheckType)))
                    {
                        if (option.IsErrorCheck(type))
                        {
                            writer.WriteLine($"    Enabled: {type}");
                        }
                    }

                    // List the ranges to which this option applies
                    int rangeCount = option.GetCountOfRange();
                    writer.WriteLine($"    Ranges count: {rangeCount}");
                    for (int r = 0; r < rangeCount; r++)
                    {
                        CellArea area = option.GetRange(r);
                        writer.WriteLine($"      {area.StartRow},{area.StartColumn} : {area.EndRow},{area.EndColumn}");
                    }
                }

                writer.WriteLine();
            }
        }

        // Save the workbook unchanged (save rule) – optional
        workbook.Save("output.xlsx");
    }
}
