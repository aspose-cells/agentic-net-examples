// Title: C# – List All Named Ranges in an Aspose.Cells Workbook and Print Their Addresses
// Description: Demonstrates how to create a workbook, define named ranges, retrieve every named range with Worksheets.GetNamedRanges(), and output each range's name and address to the console using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# list named ranges | GetNamedRanges .NET | enumerate workbook named ranges | named range address Aspose.Cells | Aspose.Cells range enumeration example
// Common Searches: Aspose.Cells get all named ranges C# | print named range addresses Aspose.Cells | list workbook named ranges .NET | How to enumerate named ranges in Aspose.Cells | Aspose.Cells GetNamedRanges usage
// Developer Intent: Retrieve every defined named range in a workbook and display its name and cell address.
// Use Cases: Validate that required named ranges exist before running calculations | Generate a quick reference sheet of all named ranges for documentation | Dynamically map named ranges to data‑processing logic at runtime
// AI Prompts: Generate C# code with Aspose.Cells that writes all named range names and addresses to a CSV file. | Show how to filter named ranges by worksheet name and output only those matches. | Provide an example that logs each named range’s details to a text file instead of the console.

using System;
using Aspose.Cells;

namespace NamedRangeEnumerationDemo
{
    // Demonstrates how to create a workbook, define named ranges, retrieve every named range with Worksheets.GetNamedRanges(), and output each range's name and address to the console using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Add sample data and named ranges for demonstration
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["A2"].PutValue("Quantity");
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);

                // Create two named ranges
                sheet.Cells.CreateRange("A1:B1").Name = "HeaderRange";
                sheet.Cells.CreateRange("A2:B2").Name = "DataRange";

                // Retrieve all pre-defined named ranges using the GetNamedRanges method (rule)
                Aspose.Cells.Range[] namedRanges = workbook.Worksheets.GetNamedRanges();

                // Output the addresses of the named ranges to the console
                if (namedRanges != null && namedRanges.Length > 0)
                {
                    Console.WriteLine($"Found {namedRanges.Length} named range(s):");
                    foreach (Aspose.Cells.Range range in namedRanges)
                    {
                        // Each Range object provides its address via the Address property
                        Console.WriteLine($"Name: {range.Name}, Address: {range.Address}");
                    }
                }
                else
                {
                    Console.WriteLine("No named ranges found in the workbook.");
                }

                // Save the workbook if needed (lifecycle rule: save)
                workbook.Save("NamedRangesDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
