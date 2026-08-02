using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // ---------- Create ----------
        // Initialize a new workbook instance
        Workbook workbook = new Workbook();

        // Set the built‑in Title property
        workbook.BuiltInDocumentProperties.Title = "Sample Document Title";

        // Save the workbook to disk (required before loading)
        string filePath = "sample.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // ---------- Load ----------
        // Load the previously saved workbook from the file
        Workbook loadedWorkbook = new Workbook(filePath);

        // Retrieve the Title built‑in property for verification
        string loadedTitle = loadedWorkbook.BuiltInDocumentProperties.Title;

        // Output the retrieved title
        Console.WriteLine("Loaded Title: " + loadedTitle);
    }
}