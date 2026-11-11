// Title: Replace {{Name}} placeholder in a named range (TemplateRange) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace placeholder text in a named range of an Excel file | How to update cells containing {{Name}} within a named range with Aspose.Cells | Iterate through cells of a named range and substitute a token using Aspose.Cells .NET | C# code to find and replace {{Name}} in TemplateRange of an XLSX workbook | Replace template placeholders in Excel named ranges with Aspose.Cells library
// Tags: replace placeholder in named range Aspose.Cells | named range cell iteration C# | Aspose.Cells string replacement Excel | template range token substitution .NET | update cells in TemplateRange Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, retrieves the named range "TemplateRange", iterates over each cell, replaces any occurrence of the "{{Name}}" placeholder with "John Doe", and saves the modified file, while handling missing files and missing named range scenarios.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the range defined by the named range "TemplateRange"
            Aspose.Cells.Range templateRange = workbook.Worksheets.GetRangeByName("TemplateRange");
            if (templateRange == null)
            {
                Console.WriteLine("Named range \"TemplateRange\" not found.");
                return;
            }

            // Iterate through each cell in the named range and replace the placeholder
            foreach (Cell cell in templateRange)
            {
                if (cell.Type == CellValueType.IsString && cell.StringValue.Contains("{{Name}}"))
                {
                    string updatedValue = cell.StringValue.Replace("{{Name}}", "John Doe");
                    cell.PutValue(updatedValue);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
