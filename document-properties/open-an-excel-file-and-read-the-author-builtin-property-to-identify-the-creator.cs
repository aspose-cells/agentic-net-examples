using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the Excel file
        string filePath = "input.xlsx";
        Workbook workbook = new Workbook(filePath);

        // Retrieve the built‑in Author property
        string author = workbook.BuiltInDocumentProperties.Author;

        // Output the author name
        Console.WriteLine("Author: " + author);
    }
}