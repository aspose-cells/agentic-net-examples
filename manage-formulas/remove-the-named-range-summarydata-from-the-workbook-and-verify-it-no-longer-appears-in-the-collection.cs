// Title: Remove a named range (SummaryData) from an Aspose.Cells workbook using C#
// Description: Demonstrates how to add a named range called "SummaryData" to a new workbook, display the named‑range count, delete the range with NameCollection.Remove, verify its absence by iterating the collection, and save the file. Includes basic error handling for robust execution.
// Keywords: Aspose.Cells remove named range C# | delete defined name SummaryData | NameCollection.Remove example | verify named range deletion | Aspose.Cells workbook cleanup | C# Aspose.Cells named range management | check named range count after removal
// Common Searches: How to delete a named range in Aspose.Cells .NET | Verify removal of a defined name in Aspose.Cells | Aspose.Cells C# remove specific named range | Count named ranges before and after deletion Aspose.Cells | Programmatically clean up named ranges with Aspose.Cells
// Developer Intent: Delete the "SummaryData" named range from a workbook and confirm that it no longer exists.
// Use Cases: Remove temporary named ranges after generating a report to keep the workbook tidy. | Ensure no leftover defined names remain before distributing a template. | Automate cleanup of obsolete named ranges when updating data sources in a CI/CD pipeline.
// AI Prompts: Generate C# code that uses Aspose.Cells to remove a named range called "SummaryData" and confirm its deletion. | Show how to iterate through a NameCollection in Aspose.Cells to verify a specific named range is absent after removal. | Explain best practices for exception handling when deleting named ranges in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a named range called "SummaryData" to a new workbook, display the named‑range count, delete the range with NameCollection.Remove, verify its absence by iterating the collection, and save the file. Includes basic error handling for robust execution.
    public class RemoveNamedRangeDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the collection of defined names
            NameCollection names = workbook.Worksheets.Names;

            // Add a named range called "SummaryData" for demonstration purposes
            int nameIndex = names.Add("SummaryData");
            names[nameIndex].RefersTo = "=Sheet1!$A$1:$B$10";

            // Display count before removal
            Console.WriteLine("Named ranges count before removal: " + names.Count);

            // Remove the named range "SummaryData"
            names.Remove("SummaryData");

            // Verify that "SummaryData" no longer exists in the collection
            bool exists = false;
            foreach (Name n in names)
            {
                if (n.Text == "SummaryData")
                {
                    exists = true;
                    break;
                }
            }

            Console.WriteLine("Does 'SummaryData' still exist? " + exists);
            Console.WriteLine("Named ranges count after removal: " + names.Count);

            // Save the workbook
            workbook.Save("RemoveSummaryDataDemo.xlsx");
        }
    }
}
