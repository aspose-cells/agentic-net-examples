// Title: How to enumerate all named ranges in an Excel workbook and write their names and addresses to the debug log using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells list all named ranges with their RefersTo addresses | iterate through NameCollection in Aspose.Cells and output to Debug.WriteLine | retrieve named range address from workbook using Aspose.Cells .NET | log Excel named range names and addresses in C# with Aspose.Cells | how to get name and address of each defined name in an Excel file using Aspose.Cells
// Tags: Aspose.Cells enumerate named ranges | C# retrieve RefersTo address Aspose.Cells | debug log Excel named ranges .NET | NameCollection iteration Aspose.Cells | list defined names workbook Aspose.Cells | extract named range addresses C#

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, accesses its NameCollection, iterates over each defined name, extracts the range's name and RefersTo address, and writes the information to the debug output using Debug.WriteLine.
class Program
{
    static void Main()
    {
        // Path to the input workbook
        string filePath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook (no password supplied)
            workbook = new Workbook(filePath);
        }
        catch (CellsException ex)
        {
            // Handle password-protected workbook
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a valid password to open it.");
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Get the collection of all named ranges in the workbook
            NameCollection names = workbook.Worksheets.Names;

            // Iterate through each named range and output its details
            foreach (Name name in names)
            {
                string rangeName = name.Text;   // The name of the range
                string address = name.RefersTo; // The address it refers to (e.g., =Sheet1!$A$1:$B$5)

                Debug.WriteLine($"Name: {rangeName}, Address: {address}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing named ranges: {ex.Message}");
        }

        // (Optional) Save the workbook if any changes were made
        // workbook.Save("output.xlsx");
    }
}
