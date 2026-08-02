// Title: C# Console App to Merge Excel Workbooks Using Aspose.Cells Combine
// Description: A minimal .NET console program that creates a source and a destination workbook, adds sample data, merges them with the Workbook.Combine method, saves the result as CombinedWorkbook.xlsx, and prints a success message.
// Keywords: Aspose.Cells console merge | Workbook.Combine C# | merge Excel files .NET | save combined workbook | C# Excel workbook merging | Aspose.Cells example | Console application Aspose.Cells
// Common Searches: how to merge two Excel workbooks in a C# console app | Aspose.Cells Workbook.Combine usage example | save merged workbook with Aspose.Cells | create console project that combines Excel files
// Developer Intent: Combine a source workbook into a destination workbook and persist the merged file from a console application using Aspose.Cells.
// Use Cases: Consolidate monthly reports and summary sheets into a single Excel file. | Automate template‑data workbook merging during build or deployment pipelines. | Generate a unified report by sequentially merging multiple generated workbooks in CI/CD.
// AI Prompts: Guide me through setting up a Visual Studio console project that uses Aspose.Cells to merge two workbooks and save the output as an Xlsx file. | Write C# code that takes an array of workbook paths, merges them with Workbook.Combine, and includes robust error handling for missing or corrupted files. | Explain how to apply an Aspose.Cells license in a console application before performing workbook merging.

using System;
using Aspose.Cells;

namespace WorkbookMergingConsoleApp
{
    // A minimal .NET console program that creates a source and a destination workbook, adds sample data, merges them with the Workbook.Combine method, saves the result as CombinedWorkbook.xlsx, and prints a success message.
    class Program
    {
        static void Main(string[] args)
        {
            // Create the source workbook and add some data
            Workbook sourceWorkbook = new Workbook(); // using Workbook() constructor
            sourceWorkbook.Worksheets[0].Cells["A1"].PutValue("Source Data");

            // Create the destination workbook (empty workbook with default format)
            Workbook destWorkbook = new Workbook(); // using Workbook() constructor
            destWorkbook.Worksheets[0].Cells["B2"].PutValue("Destination Data");

            // Merge the source workbook into the destination workbook
            destWorkbook.Combine(sourceWorkbook); // using Workbook.Combine method

            // Save the merged workbook to disk
            string outputPath = "CombinedWorkbook.xlsx";
            destWorkbook.Save(outputPath, SaveFormat.Xlsx); // using Workbook.Save method

            Console.WriteLine($"Workbooks merged successfully. Output saved to: {outputPath}");
        }
    }
}
