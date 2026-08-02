// Title: C# – List All Named Ranges in an Aspose.Cells Workbook and Output Their Addresses
// Description: Shows how to create a workbook, add sample named ranges, retrieve all defined ranges with Worksheets.GetNamedRanges(), and write each range’s name and address to the console before saving the file.
// Keywords: Aspose.Cells | C# | named ranges | list named ranges | GetNamedRanges | range address | enumerate ranges | console output | Workbook automation | Excel API
// Common Searches: Aspose.Cells GetNamedRanges C# example | list all named ranges in a workbook using Aspose | print named range addresses to console C# | retrieve named range names and addresses Aspose.Cells | enumerate named ranges Excel with Aspose.Cells
// Developer Intent: Retrieve every named range in a workbook and display its name and address.
// Use Cases: Verify that required named ranges exist before running calculations. | Generate a quick documentation snapshot of all defined ranges. | Debug complex spreadsheets by listing range definitions in the console.
// AI Prompts: Provide C# code to filter named ranges by a prefix and output the matching ones. | Show how to export named range names and addresses to a CSV file using Aspose.Cells. | Explain handling of named ranges that reference cells on different worksheets.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Shows how to create a workbook, add sample named ranges, retrieve all defined ranges with Worksheets.GetNamedRanges(), and write each range’s name and address to the console before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add sample named ranges for demonstration
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.CreateRange("A1:B2").Name = "SalesData";
            sheet.Cells.CreateRange("C3:D4").Name = "Expenses";

            // Retrieve all pre‑defined named ranges (using GetNamedRanges)
            AsposeRange[] namedRanges = workbook.Worksheets.GetNamedRanges();

            // Output each named range's name and address to the console
            if (namedRanges != null && namedRanges.Length > 0)
            {
                foreach (AsposeRange range in namedRanges)
                {
                    Console.WriteLine($"Name: {range.Name}, Address: {range.Address}");
                }
            }
            else
            {
                Console.WriteLine("No named ranges found.");
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("NamedRangesDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
