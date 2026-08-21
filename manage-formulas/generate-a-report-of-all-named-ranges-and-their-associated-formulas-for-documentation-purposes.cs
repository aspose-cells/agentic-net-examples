// Title: C# – Export All Named Ranges and Their RefersTo Formulas to a Text Report with Aspose.Cells for .NET
// Description: This example creates a workbook, defines named ranges, then iterates the NameCollection to write each range's name, RefersTo formula, and concrete cell addresses into a plain‑text report. The workbook is also saved for reference.
// Keywords: Aspose.Cells .NET named ranges | C# export named ranges to text | list RefersTo formulas Aspose.Cells | retrieve range addresses from Name.GetRanges | generate named range documentation
// Common Searches: list all named ranges with formulas using Aspose.Cells C# | how to export named range addresses to a txt file | Aspose.Cells get RefersTo property of named ranges | iterate NameCollection Aspose.Cells example | C# code to document workbook named ranges
// Developer Intent: Produce a text file that enumerates every named range in a workbook together with its RefersTo expression and the actual cell addresses.
// Use Cases: Create audit documentation that shows how cells are grouped by named ranges. | Generate user‑facing reference material for complex spreadsheets. | Validate named‑range definitions during automated testing of spreadsheet logic.
// AI Prompts: Write C# code with Aspose.Cells that reads an existing workbook and outputs all named ranges and their RefersTo strings to a CSV file. | Show how to extract each address returned by Name.GetRanges() and include it in a formatted report. | Explain how to handle named ranges that reference external sheets or formulas when building a documentation report with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// This example creates a workbook, defines named ranges, then iterates the NameCollection to write each range's name, RefersTo formula, and concrete cell addresses into a plain‑text report. The workbook is also saved for reference.
class NamedRangesReport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate some sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["B1"].PutValue(5);
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["B3"].PutValue(25);

            // Define named ranges
            int idxNumbers = workbook.Worksheets.Names.Add("Numbers");
            workbook.Worksheets.Names[idxNumbers].RefersTo = "=Sheet1!$A$1:$A$3";

            int idxValues = workbook.Worksheets.Names.Add("Values");
            workbook.Worksheets.Names[idxValues].RefersTo = "=Sheet1!$B$1:$B$3";

            // Generate a textual report of all named ranges and their formulas
            using (StreamWriter writer = new StreamWriter("NamedRangesReport.txt"))
            {
                writer.WriteLine("Named Ranges Report");
                writer.WriteLine("====================");
                writer.WriteLine();

                // Iterate through the NameCollection
                foreach (Name name in workbook.Worksheets.Names)
                {
                    writer.WriteLine($"Name       : {name.Text}");
                    writer.WriteLine($"RefersTo   : {name.RefersTo}");

                    // Retrieve the actual Range objects referenced by the name
                    Aspose.Cells.Range[] ranges = name.GetRanges();

                    if (ranges != null && ranges.Length > 0)
                    {
                        foreach (Aspose.Cells.Range range in ranges)
                        {
                            writer.WriteLine($"  Range Address : {range.Address}");
                        }
                    }
                    else
                    {
                        writer.WriteLine("  No concrete range returned.");
                    }

                    writer.WriteLine();
                }
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("SampleWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
