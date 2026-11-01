// Title: Read a cell's displayed string, replace {Name} and {Date} placeholders, and write the updated text back using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read cell string value and replace placeholders | How to substitute {Name} and {Date} tokens in an Excel cell using Aspose.Cells | Update formatted cell text programmatically with Aspose.Cells .NET | Replace placeholder tokens in Excel worksheet cell and save workbook using Aspose.Cells | Aspose.Cells get cell StringValue, modify and put back value in C#
// Tags: read cell StringValue with Aspose.Cells | replace placeholder tokens in cell value C# | write modified string to Excel cell Aspose.Cells | dynamic placeholder substitution in .NET Excel workbook | save workbook after cell text modification Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads (or creates) an Excel workbook, reads the displayed string from a target cell, replaces the {Name} and {Date} placeholders with actual values, writes the updated string back to the same cell, and saves the workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Initialize placeholder content in A1 for demonstration
                workbook.Worksheets[0].Cells["A1"].PutValue("{Name} - {Date}");
            }

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the target cell (e.g., A1)
            Cell cell = worksheet.Cells["A1"];

            // Retrieve the cell's formatted string (displayed value)
            string formattedString = cell.StringValue;

            // Replace placeholder tokens with actual values
            string updatedString = formattedString
                .Replace("{Name}", "John Doe")
                .Replace("{Date}", DateTime.Now.ToString("yyyy-MM-dd"));

            // Write the updated string back to the same cell
            cell.PutValue(updatedString);

            // Save the workbook with the changes
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
