// Title: How to retrieve watched cells (named ranges) from an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# code to list all named ranges that represent watch window cells with Aspose.Cells | Aspose.Cells .NET retrieve addresses of monitored cells in a workbook | extract watch window cell references from Excel file using Aspose.Cells API | list defined names as watch items in Aspose.Cells for .NET
// Tags: Aspose.Cells enumerate named ranges .NET | retrieve watch window cell addresses Aspose.Cells | list defined names Excel workbook C# | extract monitored cell references using Aspose.Cells | named ranges extraction Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, iterates through all defined names (named ranges), collects each RefersTo address (e.g., Sheet1!$B$2) as a potential watch‑window cell, and outputs the list while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Aspose.Cells does not expose the Watch Window via API.
            // As an alternative, retrieve all defined names (named ranges) which may represent watched cells.
            List<string> watchedCells = new List<string>();

            foreach (Name definedName in workbook.Worksheets.Names)
            {
                // The RefersTo property contains the cell address (e.g., Sheet1!$B$2)
                string refersTo = definedName.RefersTo;
                if (!string.IsNullOrEmpty(refersTo))
                {
                    watchedCells.Add(refersTo);
                }
            }

            // Output the collected addresses
            if (watchedCells.Count == 0)
            {
                Console.WriteLine("No watch items (named ranges) found.");
            }
            else
            {
                Console.WriteLine("Watched cells (named ranges):");
                foreach (string address in watchedCells)
                {
                    Console.WriteLine(address);
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
