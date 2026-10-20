// Title: Retrieve precedent cells of a formula in an Excel workbook using Aspose.Cells GetPrecedents (C#)
// Common Searches: C# Aspose.Cells get cells referenced by a formula | Using GetPrecedents to list precedent cell addresses in Excel with Aspose.Cells | Example code for retrieving precedent areas of a formula cell in Aspose.Cells .NET | How to iterate over ReferredAreaCollection returned by GetPrecedents in C# | Aspose.Cells GetPrecedents method sample for Excel file
// Tags: Aspose.Cells GetPrecedents API | enumerate precedent cells Excel .NET | ReferredAreaCollection iteration C# | retrieve formula precedents Aspose.Cells | Excel formula dependency extraction C#

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of Input.xlsx, loads it with Aspose.Cells, accesses the first worksheet, selects cell B2, calls GetPrecedents to obtain its ReferredAreaCollection, and iterates through each area to print the address of every precedent cell, handling any exceptions that may occur.
class PrecedentRetriever
{
    static void Main()
    {
        const string inputPath = "Input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Cell that contains the formula whose precedents we want
            Cell formulaCell = worksheet.Cells["B2"];

            // Retrieve the collection of precedent areas
            ReferredAreaCollection precedentAreas = formulaCell.GetPrecedents();

            // Iterate through each area and output the address of every precedent cell
            foreach (ReferredArea area in precedentAreas)
            {
                for (int row = area.StartRow; row <= area.EndRow; row++)
                {
                    for (int col = area.StartColumn; col <= area.EndColumn; col++)
                    {
                        Cell precedentCell = worksheet.Cells[row, col];
                        Console.WriteLine(precedentCell.Name);
                    }
                }
            }

            // (Optional) Save the workbook if any changes were made
            // workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
