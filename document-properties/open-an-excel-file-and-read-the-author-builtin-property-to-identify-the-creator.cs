using System;
using Aspose.Cells;

namespace AsposeCellsReadAuthor
{
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel file
            string filePath = "SampleWorkbook.xlsx";

            // Load the workbook from the file (load rule)
            Workbook workbook = new Workbook(filePath);

            // Retrieve the built‑in Author property
            string author = workbook.BuiltInDocumentProperties.Author;

            // Output the author name
            Console.WriteLine("Author: " + author);
        }
    }
}