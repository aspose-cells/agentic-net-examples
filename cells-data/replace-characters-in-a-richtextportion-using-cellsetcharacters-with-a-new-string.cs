// Title: Replace a substring in a RichTextPortion of an Excel cell using Cell.SetCharacters with Aspose.Cells for .NET (C#)
// Common Searches: how to use Cell.SetCharacters to change part of a RichTextPortion in Aspose.Cells C# | replace specific characters in an Excel cell rich text using Aspose.Cells .NET | Aspose.Cells C# example for updating a substring inside a cell with SetCharacters | modify RichTextPortion text in a worksheet cell programmatically with Aspose.Cells | C# code to replace a word in an Excel cell while preserving formatting using SetCharacters
// Tags: cell.setcharacters replace richtextportion aspose.cells | update excel cell substring c# aspose | rich text manipulation aspose.cells c# | preserve cell formatting while replacing text aspose | excel workbook save after cell text update .net

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates how to retrieve a cell's RichTextPortion, replace a target substring using the Cell.SetCharacters method to keep existing formatting, and then save the workbook, all with Aspose.Cells for .NET in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Access cell A1 and set initial text
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Hello World");

            // Replace "World" with "Aspose" in the cell's text
            string currentText = cell.StringValue;
            if (!string.IsNullOrEmpty(currentText))
            {
                string newText = currentText.Replace("World", "Aspose");
                cell.PutValue(newText);
            }

            // Determine full output path and ensure its directory exists
            string outputPath = "output.xlsx";
            string fullOutputPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullOutputPath);

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(fullOutputPath);
            Console.WriteLine($"Workbook saved successfully to '{fullOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
