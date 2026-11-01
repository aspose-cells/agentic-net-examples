// Title: Set Excel workbook author and title properties and freeze the first row using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens an existing .xlsx file, changes the BuiltInDocumentProperties Author and Title, freezes the top row of the first worksheet, and saves the updated workbook. | Show how to use the FreezePanes method to lock the first row after updating workbook metadata in a .NET console application with Aspose.Cells.
// Common Searches: C# Aspose.Cells set workbook author property programmatically | How to change Excel file title metadata with Aspose.Cells .NET | Freeze first row in an Excel worksheet using Aspose.Cells C# example | Update built-in document properties and apply freeze panes in Aspose.Cells | Aspose.Cells freeze panes after modifying workbook properties
// Tags: built-in document properties update Aspose.Cells C# | freeze first worksheet row Aspose.Cells | set Excel author title Aspose.Cells | apply FreezePanes method Aspose.Cells | modify workbook metadata programmatically Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing .xlsx workbook, updates its Author and Title built‑in document properties, freezes the first row of the first worksheet, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Update built‑in document properties
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Sales Report Q4";

            // Freeze the first row on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            // Freeze the first row (row index 1) and no columns
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
