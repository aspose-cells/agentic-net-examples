using System;
using Aspose.Cells;

class DocumentVersionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the document version property (custom version of the file)
        workbook.BuiltInDocumentProperties.DocumentVersion = "2.0";

        // Set the application version property (format "00.0000")
        workbook.BuiltInDocumentProperties.Version = "12.0000";

        // Save the workbook to disk
        string filePath = "DocumentVersionDemo.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the saved workbook to verify the properties
        Workbook loadedWorkbook = new Workbook(filePath);

        // Display the stored version properties
        Console.WriteLine("DocumentVersion: " + loadedWorkbook.BuiltInDocumentProperties.DocumentVersion);
        Console.WriteLine("Version: " + loadedWorkbook.BuiltInDocumentProperties.Version);
    }
}