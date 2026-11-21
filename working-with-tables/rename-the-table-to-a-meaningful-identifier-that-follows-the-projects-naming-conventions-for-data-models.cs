// Title: Rename an Excel ListObject table to a custom identifier using Aspose.Cells for .NET (C#)
// Common Searches: C# code to change the display name of an Excel table with Aspose.Cells | How to set a custom name for a ListObject in an .xlsx file using Aspose.Cells | Aspose.Cells .NET example for renaming the first worksheet table | Programmatically update Excel table identifier in C# with Aspose.Cells API
// Tags: Aspose.Cells ListObject DisplayName property | rename Excel table programmatically .NET | set custom table identifier Aspose.Cells | modify ListObject name in workbook C# | change Excel table display name using Aspose API

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample loads an existing workbook, checks for a ListObject table on the first worksheet, assigns a new DisplayName (e.g., "SalesData") to rename the table, and saves the workbook, demonstrating error‑handled table renaming with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Check that the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count > 0)
            {
                // Retrieve the first table
                ListObject table = sheet.ListObjects[0];

                // Rename the table to a meaningful identifier using the correct property
                table.DisplayName = "SalesData";
            }
            else
            {
                Console.WriteLine("No tables found in the worksheet.");
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
