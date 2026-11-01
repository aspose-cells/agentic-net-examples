// Title: How to set the Author built‑in document property in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set workbook Author property programmatically | update built‑in document property Author in existing .xlsx with Aspose.Cells | C# code to change Excel file author metadata using Aspose.Cells | modify Excel workbook built‑in properties Aspose.Cells .NET example | set contributor identifier as Author in Excel using Aspose.Cells library
// Tags: author built‑in property Aspose.Cells | excel metadata update C# Aspose.Cells | built‑in document properties manipulation .NET | workbook author metadata Aspose.Cells | contributor identifier Excel property Aspose

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing .xlsx (or creates a new workbook), assigns a contributor identifier to the Author built‑in document property via the BuiltInDocumentProperties indexer, and saves the workbook, demonstrating proper error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the existing Excel file
            string inputPath = "input.xlsx";

            Workbook workbook;

            // Load the workbook if the file exists; otherwise create a new workbook
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file not found: {inputPath}. A new workbook will be created.");
                workbook = new Workbook();
            }

            // Set the Author built‑in property to the contributor identifier
            string contributorId = "ContributorIdentifier"; // replace with actual ID
            // The indexer returns a BuiltInDocumentProperty; set its Value
            workbook.BuiltInDocumentProperties["Author"].Value = contributorId;

            // Save the workbook (overwrites the original or saves to a new file)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}.");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
