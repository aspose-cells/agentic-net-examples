// Title: Read the Title built‑in document property from an Excel workbook with Aspose.Cells for .NET (C#)
// Description: Loads SampleWorkbook.xlsx into an Aspose.Cells Workbook, accesses the Title built‑in document property via workbook.BuiltInDocumentProperties.Title, and prints the value to the console for verification.
// Keywords: Aspose.Cells C# read document property | Excel Title built‑in property | Workbook BuiltInDocumentProperties.Title example | load Excel file Aspose.Cells | .NET Excel metadata extraction | retrieve Excel title programmatically
// Common Searches: how to get the Title property from an Excel file using Aspose.Cells C# | C# code to read built‑in document properties with Aspose.Cells | Aspose.Cells example for retrieving workbook metadata | read Excel file title with Aspose.Cells for .NET
// Developer Intent: Extract the Title built‑in document property from an existing Excel workbook using Aspose.Cells in C#.
// Use Cases: Validate that a generated report contains the correct title before distribution. | Audit multiple workbooks to ensure consistent title metadata across files. | Log document titles during automated batch processing of Excel documents.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells and prints its Title built‑in property. | Explain how to handle missing or empty Title properties when reading Excel metadata with Aspose.Cells. | Show how to modify the Title built‑in property and save the workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsBuiltInPropertyDemo
{
    // Loads SampleWorkbook.xlsx into an Aspose.Cells Workbook, accesses the Title built‑in document property via workbook.BuiltInDocumentProperties.Title, and prints the value to the console for verification.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing Excel file
            string filePath = "SampleWorkbook.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(filePath);

            // Retrieve the Title built‑in document property
            string title = workbook.BuiltInDocumentProperties.Title;

            // Output the Title for verification
            Console.WriteLine("Document Title: " + title);
        }
    }
}
