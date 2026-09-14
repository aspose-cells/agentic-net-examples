// Title: Replace {{Name}} placeholder in a named range of an Excel template using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, retrieve the named range "TemplateRange", substitute all "{{Name}}" tokens with a given string, and persist the changes with Aspose.Cells in C#. | Loop through each cell within the specified named range and replace placeholder strings using Aspose.Cells methods. | Detect missing template file or absent named range and handle those cases while performing placeholder replacement.
// Common Searches: Aspose.Cells C# replace placeholder text in a named range of an Excel file | How to update all {{Name}} tokens in a named range using Aspose.Cells for .NET | C# code to iterate over cells in a named range and substitute values with Aspose.Cells | Saving workbook after modifying named range with Aspose.Cells | Check if named range exists before replacing text in Aspose.Cells C#
// Tags: named-range placeholder replacement Aspose.Cells C# | iterate cells in named range Aspose.Cells | update Excel template range Aspose.Cells | save modified workbook Aspose.Cells | handle missing named range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads "TemplateWorkbook.xlsx", accesses the named range "TemplateRange", replaces every occurrence of the {{Name}} placeholder with a concrete value such as "John Doe", and saves the updated workbook as "ResultWorkbook.xlsx" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "TemplateWorkbook.xlsx";
            const string resultPath = "ResultWorkbook.xlsx";

            // Ensure the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the workbook that contains the template named range
            Workbook workbook = new Workbook(templatePath);

            // Retrieve the range associated with the named range "TemplateRange"
            // GetRangeByName returns an Aspose.Cells.Range object
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName("TemplateRange");

            if (range == null)
            {
                Console.WriteLine("Named range 'TemplateRange' was not found in the workbook.");
                return;
            }

            // Replace all occurrences of the placeholder "{{Name}}" with the actual name
            foreach (Cell cell in range)
            {
                if (cell?.Value != null)
                {
                    string cellText = cell.Value.ToString();
                    if (cellText.Contains("{{Name}}"))
                    {
                        // Replace placeholder with desired value (e.g., "John Doe")
                        cell.PutValue(cellText.Replace("{{Name}}", "John Doe"));
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            // Log or display any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
