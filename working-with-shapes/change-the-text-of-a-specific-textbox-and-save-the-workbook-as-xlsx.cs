// Title: Change the text of a specific TextBox in an Excel worksheet and save the workbook as XLSX with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change TextBox content in existing XLSX workbook | Find and edit a named TextBox shape using Aspose.Cells .NET | Programmatically update Excel TextBox text and save as XLSX with Aspose.Cells | C# code to locate a TextBox by name and modify its text in Aspose.Cells | Save modified Excel file after editing TextBox using Aspose.Cells for .NET
// Tags: modify textbox text Aspose.Cells | locate shape by name Aspose.Cells | save workbook as xlsx Aspose.Cells | excel textbox manipulation .net | update shape content C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an existing XLSX workbook, searches for a TextBox named "MyTargetBox" (or falls back to the first TextBox), updates its Text property with a new string, and then saves the modified workbook as a new XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main(string[] args)
    {
        // Paths to the input and output files.
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook from the input file.
            Workbook workbook = new Workbook(inputPath);

            // Work with the first worksheet (adjust index if needed).
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one TextBox.
            if (worksheet.TextBoxes.Count == 0)
            {
                Console.WriteLine("Error: No TextBox objects found in the worksheet.");
                return;
            }

            // ------------------------------------------------------------
            // Locate the specific TextBox you want to modify.
            // ------------------------------------------------------------
            // Option 1: Find by Name (if you know the name).
            string targetName = "MyTargetBox"; // change this to the actual name if known.
            TextBox targetBox = null;

            foreach (TextBox tb in worksheet.TextBoxes)
            {
                if (tb.Name == targetName)
                {
                    targetBox = tb;
                    break;
                }
            }

            // Option 2: Fallback to the first TextBox if the named one is not found.
            if (targetBox == null)
            {
                targetBox = worksheet.TextBoxes[0];
                Console.WriteLine($"Info: TextBox named \"{targetName}\" not found. Using the first TextBox (index 0) with name \"{targetBox.Name}\".");
            }

            // ------------------------------------------------------------
            // Change the text of the identified TextBox.
            // ------------------------------------------------------------
            string newText = "This text was changed programmatically.";
            targetBox.Text = newText;
            Console.WriteLine($"Success: TextBox \"{targetBox.Name}\" text set to \"{newText}\".");

            // ------------------------------------------------------------
            // Save the workbook as XLSX.
            // ------------------------------------------------------------
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Exception occurred: {ex.Message}");
        }
    }
}
