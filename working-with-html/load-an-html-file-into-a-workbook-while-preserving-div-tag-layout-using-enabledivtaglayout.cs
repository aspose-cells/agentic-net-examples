// Title: Load an HTML file into an Aspose.Cells Workbook while preserving DIV tag layout using EnableDivTagLayout in C#
// Common Searches: C# load html into workbook with EnableDivTagLayout option | Aspose.Cells preserve div layout when converting html to xlsx | How to use HtmlLoadOptions.EnableDivTagLayout in Aspose.Cells C# | Convert HTML file to Excel while keeping div structure using Aspose.Cells | EnableDivTagLayout flag example for Aspose.Cells HTML import
// Tags: EnableDivTagLayout HtmlLoadOptions | Aspose.Cells HTML to XLSX conversion | preserve div layout Excel export | C# load HTML workbook Aspose.Cells | HTML import with layout retention Aspose

using Aspose.Cells;
using System;
using System.IO;

// The example checks that the source HTML file exists, loads it into an Aspose.Cells Workbook (optionally using HtmlLoadOptions.EnableDivTagLayout to retain DIV tag layout), saves the workbook as an XLSX file, and outputs any errors that occur during the process.
class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFilePath = "input.html";

        // Verify that the input HTML file exists
        if (!File.Exists(htmlFilePath))
        {
            Console.WriteLine($"Error: The file \"{htmlFilePath}\" was not found.");
            return;
        }

        try
        {
            // Load the HTML file into a workbook
            Workbook workbook = new Workbook(htmlFilePath);

            // Save the workbook to an Excel file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading or saving
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
