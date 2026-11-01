// Title: Add a bold RichTextPortion to an existing cell using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add bold text to part of a cell | How to use Characters method for partial formatting in Aspose.Cells | Append rich text with bold formatting to an Excel cell using Aspose.Cells .NET | Set mixed formatting (plain and bold) in a single cell with Aspose.Cells C# | Create a workbook and apply bold style to a substring of cell A1 in Aspose.Cells
// Tags: add bold rich text portion Aspose.Cells | partial cell formatting using Characters method | mixed text styling Excel cell Aspose.Cells C# | append styled substring to worksheet cell | save workbook with formatted text Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, writes plain text to cell A1, appends a bold segment, and uses the Characters(startIndex, length).Font.IsBold property to apply bold formatting only to the new portion before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Target cell that currently contains plain text
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("Plain text "); // set initial plain text

                // Retrieve the existing plain text
                string existingText = cell.StringValue;

                // Append bold portion text
                string boldText = "Bold portion";
                cell.PutValue(existingText + boldText);

                // Apply bold formatting to the newly added portion
                int startIndex = existingText.Length;
                // Use the Characters method (start index, length) to get the range
                cell.Characters(startIndex, boldText.Length).Font.IsBold = true;

                // Define output file path
                string outputPath = "Output.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
