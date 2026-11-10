// Title: Export defined names and their RefersTo formulas to an audit worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells list all defined names and their formulas in a new worksheet | how to create an audit sheet of named ranges using Aspose.Cells .NET | retrieve RefersTo address of each named range with Aspose.Cells in C# | export Excel named ranges to a separate sheet programmatically with Aspose.Cells | iterate NameCollection to write named range details to workbook using Aspose.Cells
// Tags: export defined names to worksheet Aspose.Cells | iterate NameCollection RefersTo C# | add audit worksheet for named ranges Aspose.Cells | list named ranges with formulas .NET | save workbook with additional audit sheet Aspose.Cells | extract named range addresses programmatically C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, adds a new worksheet named "NamedRangesAudit", writes column headers, iterates through the workbook's NameCollection, and records each defined name's Text and RefersTo address into the sheet before saving the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of named ranges in the workbook
            NameCollection namedRanges = workbook.Worksheets.Names;

            // Add a new worksheet that will contain the audit information
            int auditSheetIndex = workbook.Worksheets.Add();
            Worksheet auditSheet = workbook.Worksheets[auditSheetIndex];
            auditSheet.Name = "NamedRangesAudit";

            // Write header titles
            auditSheet.Cells[0, 0].PutValue("Name");
            auditSheet.Cells[0, 1].PutValue("Refers To");

            // Populate the audit sheet with each named range and its details
            int row = 1;
            foreach (Name name in namedRanges)
            {
                // Name of the defined range
                auditSheet.Cells[row, 0].PutValue(name.Text);

                // The address or range the name refers to (e.g., Sheet1!$A$1:$B$2)
                auditSheet.Cells[row, 1].PutValue(name.RefersTo);

                row++;
            }

            // Save the workbook with the audit sheet added
            workbook.Save(outputPath);
            Console.WriteLine($"Audit workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
