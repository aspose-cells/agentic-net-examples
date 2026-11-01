// Title: Unhide columns H‑J, set a 50‑point width, and export the workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells hide columns then unhide specific range before PDF export | set column width in points with Aspose.Cells .NET | export Excel to PDF after adjusting column visibility using Aspose.Cells | how to make columns H to J visible and set width in Aspose.Cells C# | Aspose.Cells column width 50 points conversion to PDF example
// Tags: Aspose.Cells column unhide range C# | Aspose.Cells column width setting C# | Aspose.Cells generate PDF from worksheet | worksheet column visibility manipulation Aspose.Cells | C# set column width characters Aspose.Cells | Aspose.Cells save workbook as PDF

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads an existing Excel file, unhides columns H through J, sets each column's width to approximately 50 points (7.5 characters), and then saves the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Unhide columns H (index 7) through J (index 9) and set their width to ~7.5 characters
            for (int col = 7; col <= 9; col++)
            {
                // Ensure the column is visible
                sheet.Cells.Columns[col].IsHidden = false;

                // Set column width (character width; 7.5 approximates 50 points)
                sheet.Cells.SetColumnWidth(col, 7.5);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
