// Title: Load an Excel workbook, set author and title properties, modify a cell, and save as XLSX using Aspose.Cells for .NET
// Common Searches: how to open an existing Excel file or create a new workbook when the file is missing with Aspose.Cells C# | setting built‑in document properties such as Author and Title in an Excel workbook using Aspose.Cells .NET | updating a cell value and saving the workbook to XLSX format with Aspose.Cells C# example | Aspose.Cells C# code for handling missing input.xlsx and creating a default workbook | preserving document properties while saving an Excel file to XLSX using Aspose.Cells
// Tags: open or create workbook Aspose.Cells C# | set built‑in document properties Aspose.Cells | modify cell value Aspose.Cells | save workbook as XLSX Aspose.Cells | handle missing input file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample opens 'input.xlsx' if it exists or creates a new workbook otherwise, sets the Author and Title built‑in document properties, writes "Hello, Aspose!" to cell A1 of the first worksheet, and saves the result as 'output.xlsx' in XLSX format, with exception handling for any errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; otherwise create a new workbook
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Set built‑in document properties
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Modified Workbook";

            // Example modification: change a cell value in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Hello, Aspose!");

            // Save the workbook to XLSX format, preserving all changes
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
