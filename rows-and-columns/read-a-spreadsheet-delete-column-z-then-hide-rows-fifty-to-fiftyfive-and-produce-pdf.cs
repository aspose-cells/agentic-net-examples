// Title: Remove column Z, hide rows 50‑55, and export the worksheet as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, removes column Z, hides rows 50 through 55, and saves the result as a PDF. | Generate a C# example using Aspose.Cells to delete a specific column, set a range of rows to hidden, and convert the modified workbook to PDF.
// Common Searches: Aspose.Cells C# delete column Z and conceal rows 50 through 55 before PDF export | How to hide a range of rows in Aspose.Cells and then save the sheet as PDF | C# code to remove a column and hide rows in an Excel workbook using Aspose.Cells | Export modified worksheet to PDF after deleting a column with Aspose.Cells .NET | Aspose.Cells example for concealing rows 50‑55 in C#
// Tags: column Z removal Aspose.Cells C# | rows 50-55 hidden state Aspose.Cells C# | PDF conversion after worksheet modification Aspose.Cells | Aspose.Cells delete column and hide rows workflow | export workbook to PDF Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The C# program loads "input.xlsx", deletes column Z (index 25), hides rows 50‑55 (indices 49‑54), and saves the updated workbook as "output.pdf" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing spreadsheet
            var workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            var worksheet = workbook.Worksheets[0];

            // Delete column Z (zero‑based index 25)
            worksheet.Cells.DeleteColumn(25);

            // Hide rows 50 to 55 (zero‑based indices 49‑54)
            for (int rowIndex = 49; rowIndex <= 54; rowIndex++)
            {
                // Use IsHidden property to hide the row
                worksheet.Cells.Rows[rowIndex].IsHidden = true;
            }

            // Save the modified workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
