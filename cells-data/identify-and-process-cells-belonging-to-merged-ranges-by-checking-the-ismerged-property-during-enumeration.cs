// Title: Enumerate and list merged cells in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to find all merged cells in an Excel file with Aspose.Cells | Aspose.Cells iterate through cells and detect merged ranges in .NET | How to output addresses and values of merged cells using Aspose.Cells library | Reading a password‑protected Excel workbook and listing merged cells in C# | Using Cell.IsMerged property to identify merged cells in Aspose.Cells
// Tags: enumerate merged cells Aspose.Cells .NET | Cell.IsMerged detection C# | list merged ranges Aspose.Cells workbook | process password protected Excel Aspose.Cells | save workbook after merged cell enumeration

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook (prompting for a password when needed), iterates through each worksheet and every cell, checks the Cell.IsMerged flag, writes each merged cell's address and value to the console, ensures the output directory exists, and saves the workbook to a new file.
class MergedCellsProcessor
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Try to load the workbook without a password
            workbook = new Workbook(inputPath);
        }
        catch (CellsException ex) when (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            // Workbook is password‑protected; ask the user for the password
            Console.Write("Enter password for the workbook: ");
            string password = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(password))
            {
                Console.WriteLine("No password entered. Cannot open the workbook.");
                return;
            }

            try
            {
                LoadOptions loadOptions = new LoadOptions { Password = password };
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook with password: {loadEx.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        try
        {
            // Iterate through each worksheet and its cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                foreach (Cell cell in cells)
                {
                    if (cell.IsMerged)
                    {
                        // Output merged cell address and value
                        Console.WriteLine($"Merged Cell: {cell.Name} (Value: {cell.Value})");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (no modifications made in this example)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during processing: {ex.Message}");
        }
    }
}
