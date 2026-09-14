// Title: Extract all named ranges from an Excel .xlsx workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, iterates through Workbook.Worksheets.Names, and prints each named range. | Show a try‑catch example that loads a workbook, gathers the Text property of every Name object into a list, and outputs the list to the console. | Demonstrate how to verify the existence of an Excel file before extracting its defined names with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# how to list all defined names in an Excel file | C# retrieve named ranges from .xlsx using Aspose.Cells library | example code to enumerate workbook named ranges with Aspose.Cells .NET | read Excel named ranges programmatically in C# Aspose.Cells | list workbook names collection Aspose.Cells C# tutorial
// Tags: Aspose.Cells enumerate named ranges | C# extract Excel defined names | list workbook names Aspose.Cells | read .xlsx named ranges .NET | Aspose.Cells named range extraction

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads an existing .xlsx workbook with Aspose.Cells, checks that the file exists, iterates through the Workbook.Worksheets.Names collection, collects each Name's Text into a list, and writes the defined (named) ranges to the console while handling errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Collect defined (named) ranges from the workbook
            List<string> definedNames = new List<string>();
            foreach (Name name in workbook.Worksheets.Names)
            {
                definedNames.Add(name.Text);
            }

            // Output the extracted defined names
            foreach (string name in definedNames)
            {
                Console.WriteLine(name);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
