using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook from disk
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Modify built‑in document properties
        workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
        workbook.BuiltInDocumentProperties["Title"].Value = "Modified Workbook";

        // Add a custom document property
        workbook.CustomDocumentProperties.Add("Reviewed", true);

        // Change the default style (font name and size)
        workbook.DefaultStyle.Font.Name = "Calibri";
        workbook.DefaultStyle.Font.Size = 11;

        // Example cell modification
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Modified Content");

        // Save the workbook to XLSX format, preserving all changes
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}