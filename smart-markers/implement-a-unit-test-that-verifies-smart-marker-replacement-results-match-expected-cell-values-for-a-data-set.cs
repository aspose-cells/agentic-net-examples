// Title: C# unit test that verifies Aspose.Cells smart marker expansion produces correct cell values
// Common Searches: how to write a unit test for Aspose.Cells smart markers in C# | verify WorkbookDesigner smart marker output using NUnit | assert cell values after processing smart markers with Aspose.Cells | C# test for smart marker row expansion in an Aspose.Cells workbook
// Tags: Aspose.Cells smart marker validation | C# WorkbookDesigner expansion test | cell value assertion Aspose.Cells | smart marker template verification C# | Aspose.Cells unit testing patterns

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Simple POCO class used as data source for smart markers
    // The example creates a workbook with smart markers for Name and Score, processes a List<Person> data source via WorkbookDesigner, and then uses a helper method to assert that each expanded cell (A2‑B4) contains the expected values, throwing an exception on any mismatch.
    public class Person
    {
        public string Name { get; set; }
        public int Score { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // 2. Get the first worksheet and set up the template with smart markers
                Worksheet sheet = workbook.Worksheets[0];
                // Header row
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Score");
                // Smart marker row – will be expanded for each data item
                sheet.Cells["A2"].PutValue("&=Data.Name");
                sheet.Cells["B2"].PutValue("&=Data.Score");

                // 3. Prepare the data source
                List<Person> data = new List<Person>
                {
                    new Person { Name = "Alice", Score = 85 },
                    new Person { Name = "Bob",   Score = 92 },
                    new Person { Name = "Carol", Score = 78 }
                };

                // 4. Process smart markers using WorkbookDesigner (lifecycle: load/process)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Data", data);
                designer.Process(); // expands the smart marker row and fills values

                // Helper to read cell value as string for comparison
                string GetCellText(string address) => sheet.Cells[address].StringValue;

                // 5. Verify that the resulting cells contain the expected values
                VerifyCell("A2", "Alice", GetCellText);
                VerifyCell("B2", "85",    GetCellText);
                VerifyCell("A3", "Bob",   GetCellText);
                VerifyCell("B3", "92",    GetCellText);
                VerifyCell("A4", "Carol", GetCellText);
                VerifyCell("B4", "78",    GetCellText);

                Console.WriteLine("All smart marker replacements verified successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static void VerifyCell(string address, string expected, Func<string, string> getCellText)
        {
            string actual = getCellText(address);
            if (!string.Equals(actual, expected, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    $"Verification failed at {address}: expected '{expected}', got '{actual}'.");
            }
        }
    }
}
