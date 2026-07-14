using System;
using Aspose.Cells;

namespace AsposeCellsReadAuthor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file
            string filePath = "Sample.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Retrieve the built‑in Author property
            string author = workbook.BuiltInDocumentProperties.Author;

            // Output the author name
            Console.WriteLine("Author: " + author);
        }
    }
}