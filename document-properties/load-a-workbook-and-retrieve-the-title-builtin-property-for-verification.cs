using System;
using Aspose.Cells;

namespace AsposeCellsBuiltInPropertyDemo
{
    class Program
    {
        static void Main()
        {
            // Path for the temporary workbook file
            string filePath = "BuiltInTitleDemo.xlsx";

            // ---------- Create a new workbook and set the Title property ----------
            Workbook workbook = new Workbook(); // create
            workbook.BuiltInDocumentProperties.Title = "Sample Document Title";

            // Save the workbook to disk (lifecycle rule: use Save)
            workbook.Save(filePath, SaveFormat.Xlsx);

            // ---------- Load the saved workbook ----------
            Workbook loadedWorkbook = new Workbook(filePath); // load

            // Retrieve the Title built‑in property for verification
            string title = loadedWorkbook.BuiltInDocumentProperties.Title;

            // Output the retrieved title
            Console.WriteLine("Loaded Document Title: " + title);
        }
    }
}