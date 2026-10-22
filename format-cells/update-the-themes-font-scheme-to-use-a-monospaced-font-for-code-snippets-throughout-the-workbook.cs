// Title: Set the default workbook font to a monospaced type (Consolas) using Aspose.Cells for .NET (C#)
// Common Searches: how to change the default font of an Excel workbook to Consolas with Aspose.Cells in C# | Aspose.Cells C# set monospaced font for all cells in a workbook | apply a global monospaced style to an existing Excel file using Aspose.Cells .NET | C# code to set workbook default style font to Consolas with Aspose.Cells | update Excel theme font to a code‑friendly monospaced font using Aspose.Cells
// Tags: Aspose.Cells set default style font | C# change workbook default font | apply monospaced font to Excel workbook | Consolas default font Aspose.Cells | global font update Excel .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook (or creates a new one), sets the workbook's DefaultStyle font to the monospaced font Consolas, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // If the Aspose.Cells version supports themes, you could modify the theme here.
            // As a fallback, set the default font to a monospaced font (e.g., Consolas).
            workbook.DefaultStyle.Font.Name = "Consolas";

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
