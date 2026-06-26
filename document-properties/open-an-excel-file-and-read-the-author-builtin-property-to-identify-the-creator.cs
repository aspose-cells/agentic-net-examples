using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(filePath);

        // Retrieve the built‑in Author property
        string author = workbook.BuiltInDocumentProperties.Author;

        // Display the author name
        Console.WriteLine("Author: " + author);
    }
}