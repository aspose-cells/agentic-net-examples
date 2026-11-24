// Title: Export named ranges and their formulas to a summary worksheet using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export defined names with formulas to a new sheet | list all named ranges and their formulas in an Excel workbook with Aspose.Cells | create a summary worksheet of named ranges in a .xlsx file using Aspose.Cells | retrieve the formula of a named range in Aspose.Cells .NET | write name, refers‑to, and formula of each defined name to a separate worksheet in C# | iterate workbook.Names and generate a named‑range report with Aspose.Cells
// Tags: Aspose.Cells export named ranges to worksheet | C# retrieve named range formulas Aspose.Cells | Aspose.Cells iterate workbook names | Excel .xlsx named range summary generation | Aspose.Cells auto‑fit columns after export | Aspose.Cells GetRange formula extraction

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, adds a new worksheet called "NamedRangesSummary", writes headers (Name, Refers To, Formula), and then iterates through all defined names in the workbook. For each name it records the name text, the RefersTo expression, and determines the associated formula—either directly from RefersTo when it starts with '=', or by fetching the first cell's formula from the referenced range. The collected information is written to the summary sheet, columns are auto‑fitted for readability, and the workbook is saved to the specified output file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;
        try
        {
            // Load the existing workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Add a new worksheet for the summary
        int summaryIndex = workbook.Worksheets.Add();
        Worksheet summarySheet = workbook.Worksheets[summaryIndex];
        summarySheet.Name = "NamedRangesSummary";

        // Write header row
        summarySheet.Cells[0, 0].PutValue("Name");
        summarySheet.Cells[0, 1].PutValue("Refers To");
        summarySheet.Cells[0, 2].PutValue("Formula");

        int row = 1; // start writing data from the second row

        // Iterate through all defined names in the workbook
        foreach (Name name in workbook.Worksheets.Names)
        {
            // Name of the defined range
            summarySheet.Cells[row, 0].PutValue(name.Text);

            // The address or expression the name refers to
            summarySheet.Cells[row, 1].PutValue(name.RefersTo);

            // Determine the formula associated with the name, if any
            string formula = string.Empty;

            // In older Aspose.Cells versions Name does not expose IsFormula/Formula,
            // so we infer a formula when RefersTo starts with '='
            if (!string.IsNullOrEmpty(name.RefersTo) && name.RefersTo.StartsWith("="))
            {
                formula = name.RefersTo; // keep the leading '=' for clarity
            }
            else
            {
                // The name refers to a range; attempt to retrieve a formula from the first cell
                try
                {
                    Aspose.Cells.Range range = name.GetRange();
                    if (range != null && range.RowCount > 0 && range.ColumnCount > 0)
                    {
                        Cell firstCell = range[0, 0];
                        if (!string.IsNullOrEmpty(firstCell.Formula))
                        {
                            formula = firstCell.Formula;
                        }
                    }
                }
                catch
                {
                    // If GetRange fails (e.g., external reference), leave formula empty
                }
            }

            summarySheet.Cells[row, 2].PutValue(formula);
            row++;
        }

        // Adjust column widths for readability
        summarySheet.AutoFitColumns();

        // Save the workbook with the new summary sheet
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
