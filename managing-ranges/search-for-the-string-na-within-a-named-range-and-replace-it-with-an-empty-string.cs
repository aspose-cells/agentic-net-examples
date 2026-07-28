// Title: C# – Replace "#N/A" in a Named Range using Aspose.Cells
// Description: Demonstrates how to create a workbook, define a named range (A1:A4), iterate its cells, detect the literal "#N/A" text, replace each occurrence with an empty string, and save the file as an XLSX document with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# replace #N/A | named range replace text | Aspose.Cells iterate range cells | remove #N/A from Excel | C# Excel placeholder cleanup
// Common Searches: replace #N/A in named range Aspose.Cells | Aspose.Cells C# find and replace text in range | how to clear error strings in Excel using Aspose | C# iterate named range cells Aspose.Cells
// Developer Intent: Replace every "#N/A" string inside a specific named range with an empty value using Aspose.Cells for .NET.
// Use Cases: Sanitize imported spreadsheets by removing placeholder error strings before calculations. | Prepare data for export or reporting by clearing "#N/A" entries from defined ranges. | Automate data cleaning in batch processes that rely on named ranges.
// AI Prompts: Write C# code with Aspose.Cells that searches a named range for "#N/A" and replaces each occurrence with an empty string, then saves the workbook. | Show how to retrieve a Range object from a workbook name and iterate its cells to modify values in Aspose.Cells for .NET. | Provide an example of defining a named range, looping through its cells, and performing a conditional replace operation using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, define a named range (A1:A4), iterate its cells, detect the literal "#N/A" text, replace each occurrence with an empty string, and save the file as an XLSX document with Aspose.Cells for .NET.
    public class ReplaceNaInNamedRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with data and "#N/A" placeholders
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["A2"].PutValue("#N/A");
                sheet.Cells["A3"].PutValue("Data");
                sheet.Cells["A4"].PutValue("#N/A");

                // Define a named range covering A1:A4
                int nameIdx = workbook.Worksheets.Names.Add("MyRange");
                Name myRangeName = workbook.Worksheets.Names[nameIdx];
                myRangeName.RefersTo = $"={sheet.Name}!$A$1:$A$4";

                // Get the actual range object
                Aspose.Cells.Range range = myRangeName.GetRange();

                // Replace "#N/A" values with empty strings
                foreach (Cell cell in range)
                {
                    if (cell.StringValue == "#N/A")
                    {
                        cell.PutValue(string.Empty);
                    }
                }

                // Save the workbook
                string outputPath = "ReplaceNaInNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    class Program
    {
        static void Main(string[] args)
        {
            ReplaceNaInNamedRange.Run();
        }
    }
}
