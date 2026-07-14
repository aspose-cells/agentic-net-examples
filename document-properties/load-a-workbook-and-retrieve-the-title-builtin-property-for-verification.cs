using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and set its Title built‑in property
        Workbook workbook = new Workbook();
        workbook.BuiltInDocumentProperties.Title = "Sample Document Title";

        // Save the workbook to a file
        string filePath = "SampleTitle.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the workbook from the saved file
        Workbook loadedWorkbook = new Workbook(filePath);

        // Retrieve the Title built‑in property for verification
        string title = loadedWorkbook.BuiltInDocumentProperties.Title;

        // Display the retrieved title
        Console.WriteLine("Loaded Title: " + title);
    }
}