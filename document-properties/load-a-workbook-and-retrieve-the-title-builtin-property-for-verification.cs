using System;
using Aspose.Cells;

namespace AsposeCellsBuiltInPropertyDemo
{
    class Program
    {
        static void Main()
        {
            // Define the file path for the workbook
            string filePath = "BuiltInTitleDemo.xlsx";

            // -------------------------------------------------
            // Create a new workbook and set the Title property
            // -------------------------------------------------
            Workbook workbook = new Workbook(); // create a new workbook
            workbook.BuiltInDocumentProperties.Title = "Sample Document Title"; // set Title

            // Save the workbook to disk
            workbook.Save(filePath, SaveFormat.Xlsx);

            // -------------------------------------------------
            // Load the workbook from the saved file
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook(filePath); // load the workbook

            // Retrieve the Title built‑in property for verification
            string retrievedTitle = loadedWorkbook.BuiltInDocumentProperties.Title;

            // Output the retrieved Title to the console
            Console.WriteLine("Retrieved Title: " + retrievedTitle);
        }
    }
}