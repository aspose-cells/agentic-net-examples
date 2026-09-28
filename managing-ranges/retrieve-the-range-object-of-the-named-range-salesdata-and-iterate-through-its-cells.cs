// Title: Retrieve the 'SalesData' named range and iterate its cells using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook, obtains the Range object for the named range "SalesData" with Aspose.Cells, and prints each cell's address and value. | Show how to verify a named range exists in a workbook and then loop through all its cells using Aspose.Cells in a .NET console application.
// Common Searches: aspocells c# get range object for named range salesdata | c# iterate over cells in a named range using Aspose.Cells | how to read values from a named range in an Excel file with Aspose.Cells .NET | aspocells check if named range exists before iterating cells | retrieve named range and print cell addresses with Aspose.Cells C#
// Tags: Aspose.Cells get named range | Aspose.Cells iterate cells in range | Aspose.Cells retrieve Range object | Aspose.Cells read named range values | Aspose.Cells .NET workbook named range

using Aspose.Cells;
using System;
using System.IO;

// The example loads "input.xlsx", confirms the presence of the named range "SalesData", obtains its Range object via workbook.Worksheets.Names, and iterates through each cell, outputting the cell name and its value to the console, with error handling for missing files or ranges.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Retrieve the named range "SalesData"
            Name salesRangeName = workbook.Worksheets.Names["SalesData"];
            if (salesRangeName == null)
            {
                Console.WriteLine("Named range 'SalesData' not found.");
                return;
            }

            // Get the Range object that the name refers to
            Aspose.Cells.Range salesRange = salesRangeName.GetRange();

            // Iterate through each cell in the range
            foreach (Cell cell in salesRange)
            {
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
