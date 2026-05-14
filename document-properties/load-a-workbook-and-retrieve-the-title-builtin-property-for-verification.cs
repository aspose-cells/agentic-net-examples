using System;
using Aspose.Cells;

namespace AsposeCellsBuiltInPropertyDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded.
            // Ensure that the file exists at this location.
            string filePath = "input.xlsx";

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(filePath);

            // Retrieve the Title built‑in document property.
            string title = workbook.BuiltInDocumentProperties.Title;

            // Output the retrieved title for verification.
            Console.WriteLine("Document Title: " + title);
        }
    }
}