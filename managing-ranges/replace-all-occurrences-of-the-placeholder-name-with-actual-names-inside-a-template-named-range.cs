// Title: Replace {{Name}} placeholders in a named range of an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to replace {{Name}} placeholder inside a named range with Aspose.Cells | How to update template range cells in an Excel file using Aspose.Cells .NET | Aspose.Cells replace text in a specific named range programmatically | Iterate over cells of a named range and modify values with Aspose.Cells in C# | Replace placeholder values in an Excel template using the Aspose.Cells workbook API
// Tags: Aspose.Cells named range text replacement | C# iterate cells in Excel named range Aspose.Cells | Excel template placeholder update using Aspose.Cells | modify cell values in named range .NET | Aspose.Cells workbook placeholder substitution

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, retrieves the named range "TemplateRange", iterates through each cell within that range, and replaces any occurrence of the "{{Name}}" placeholder with "John Doe". It then saves the modified workbook, handling file existence checks and exceptions.
class Program
{
    static void Main()
    {
        // Define input and output file paths (replace placeholders with actual paths)
        string inputPath = "{InputPath}";
        string outputPath = "{OutputPath}";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Retrieve the named range "TemplateRange"
            var range = workbook.Worksheets.GetRangeByName("TemplateRange");
            if (range != null)
            {
                // Access the cells of the worksheet where the named range resides
                var cells = range.Worksheet.Cells;

                // Determine the start and end rows/columns of the named range
                int startRow = range.FirstRow;
                int endRow = range.FirstRow + range.RowCount - 1;
                int startColumn = range.FirstColumn;
                int endColumn = range.FirstColumn + range.ColumnCount - 1;

                // Iterate through each cell inside the named range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        // Replace placeholder "{{Name}}" with actual value
                        if (cell.Value != null && cell.StringValue.Contains("{{Name}}"))
                        {
                            cell.PutValue(cell.StringValue.Replace("{{Name}}", "John Doe"));
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
