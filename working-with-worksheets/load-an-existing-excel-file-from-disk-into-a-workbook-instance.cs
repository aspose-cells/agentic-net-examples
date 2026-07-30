// Title: Load an Existing Excel Workbook and Retrieve Worksheet Details with Aspose.Cells (C#)
// Description: Shows how to create a Workbook from a local .xlsx file using Aspose.Cells for .NET, access the first worksheet, and output the total number of sheets and the first sheet's name to the console.
// Keywords: Aspose.Cells load workbook C# | open Excel file Aspose.Cells | read worksheet count Aspose.Cells | first worksheet name C# | Aspose.Cells .NET example | load Excel from file path | C# console Aspose.Cells
// Common Searches: C# Aspose.Cells open existing .xlsx | How to get number of sheets after loading workbook Aspose.Cells | Retrieve first worksheet name using Aspose.Cells C# | Aspose.Cells load workbook from disk example | Read Excel workbook metadata with Aspose.Cells
// Developer Intent: Open a local Excel file as a Workbook and extract basic sheet metadata.
// Use Cases: Initialize a template workbook for data population before saving. | Iterate through sheet names to drive dynamic report generation. | Validate workbook structure (sheet count, names) prior to import operations.
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells and prints all worksheet names. | Show how to catch and handle FileNotFoundException when creating a Workbook from a path. | Demonstrate loading an Excel workbook from a MemoryStream using Aspose.Cells in C#. | Explain how to load a password‑protected workbook with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    // Shows how to create a Workbook from a local .xlsx file using Aspose.Cells for .NET, access the first worksheet, and output the total number of sheets and the first sheet's name to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing Excel file on disk
            string filePath = "input.xlsx";

            // Load the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet firstSheet = workbook.Worksheets[0];

            // Output basic information about the loaded workbook
            Console.WriteLine($"Workbook loaded successfully.");
            Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
            Console.WriteLine($"First worksheet name: {firstSheet.Name}");
        }
    }
}
