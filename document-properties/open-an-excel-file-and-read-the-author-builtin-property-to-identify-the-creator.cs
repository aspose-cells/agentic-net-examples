// Title: Read the Author built-in document property from an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read Author property from .xlsx file | how to retrieve Excel file creator using Aspose.Cells API | get built-in document properties author in Aspose.Cells C# | load workbook and access author metadata with Aspose.Cells | Aspose.Cells example for reading Excel author name in .NET
// Tags: Aspose.Cells read built-in document property | C# get Excel author metadata | Workbook BuiltInDocumentProperties Author example | extract creator name from .xlsx using Aspose | load Excel workbook and read author property

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, reads the Author built-in document property via workbook.BuiltInDocumentProperties.Author, outputs the author name, and handles any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(filePath);

            // Retrieve the Author built‑in document property
            string author = workbook.BuiltInDocumentProperties.Author;

            // Display the author name
            Console.WriteLine("Author: " + author);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
