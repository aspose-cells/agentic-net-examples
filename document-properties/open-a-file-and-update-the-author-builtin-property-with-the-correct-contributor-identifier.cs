// Title: How to update the Author built-in property of an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set built-in Author property in existing Excel file | programmatically change Excel workbook Author metadata using Aspose.Cells .NET | update built-in document property Author in .xlsx with Aspose.Cells API | C# example for modifying Excel Author property via Aspose.Cells Workbook | set contributor ID as Author in Excel using Aspose.Cells for .NET
// Tags: Aspose.Cells set built-in Author property | C# update Excel document property Author | modify workbook built-in properties Aspose.Cells | write ContributorID to Excel Author metadata | save workbook after property modification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample checks for the input.xlsx file, loads it with Aspose.Cells, retrieves the built-in Author property, assigns the value "ContributorID123", and saves the workbook as output.xlsx, handling any runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Update the built‑in Author property
                var authorProperty = workbook.BuiltInDocumentProperties["Author"];
                if (authorProperty != null)
                {
                    authorProperty.Value = "ContributorID123";
                }

                // Save the workbook (overwrites or creates a new file)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
