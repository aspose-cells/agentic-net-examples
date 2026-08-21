// Title: C# Console App to Merge Excel Workbooks Using Aspose.Cells Combine
// Description: A step‑by‑step console program that creates sample .xlsx files (if missing), loads them with Aspose.Cells, merges them via the Workbook.Combine method, and saves the result as CombinedWorkbook.xlsx.
// Keywords: Aspose.Cells combine workbooks | C# merge Excel files | Workbook.Combine example | create sample workbook Aspose.Cells | save merged workbook C#
// Common Searches: merge two Excel files with Aspose.Cells console app | C# code to combine multiple workbooks using Aspose.Cells | how to create a sample Excel workbook programmatically C#
// Developer Intent: Build a .NET console application that consolidates several Excel workbooks into a single file using Aspose.Cells.
// Use Cases: Consolidate daily reports from multiple sources into one workbook for executive review. | Generate placeholder workbooks automatically when source files are absent before merging. | Integrate workbook merging into automated build or CI pipelines with command‑line arguments.
// AI Prompts: Write a C# console program that accepts a list of Excel file paths and merges them into one workbook using Aspose.Cells. | Show how to preserve original worksheet names and resolve duplicates when combining workbooks with Aspose.Cells. | Add robust error handling for missing files, unsupported formats, and permission issues while using Workbook.Combine.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookMergeApp
{
    // A step‑by‑step console program that creates sample .xlsx files (if missing), loads them with Aspose.Cells, merges them via the Workbook.Combine method, and saves the result as CombinedWorkbook.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the source workbooks (you can change these or pass via command‑line arguments)
            string sourcePath1 = "Source1.xlsx";
            string sourcePath2 = "Source2.xlsx";

            // Create sample source files if they do not exist (for demonstration purposes)
            CreateSampleWorkbook(sourcePath1, "Data from source 1");
            CreateSampleWorkbook(sourcePath2, "Data from source 2");

            // Load the source workbooks using the Workbook(string) constructor
            Workbook sourceWorkbook1 = new Workbook(sourcePath1);
            Workbook sourceWorkbook2 = new Workbook(sourcePath2);

            // Create an empty destination workbook using the default constructor
            Workbook destinationWorkbook = new Workbook();

            // Combine the source workbooks into the destination workbook
            destinationWorkbook.Combine(sourceWorkbook1);
            destinationWorkbook.Combine(sourceWorkbook2);

            // Save the combined workbook using the Save(string, SaveFormat) method
            string outputPath = "CombinedWorkbook.xlsx";
            destinationWorkbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbooks merged successfully. Output saved to: {outputPath}");
        }

        // Helper method to create a simple workbook with a single cell value
        static void CreateSampleWorkbook(string filePath, string cellValue)
        {
            if (File.Exists(filePath))
                return;

            // Create a new workbook, put a value in A1, and save it
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue(cellValue);
            wb.Save(filePath, SaveFormat.Xlsx);
        }
    }
}
