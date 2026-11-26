// Title: C# example: Extract text from every TextBox shape in an Excel workbook and save to a CSV file using Aspose.Cells
// Common Searches: C# Aspose.Cells extract all textbox strings from an Excel file and write to CSV | How to iterate over TextBoxCollection in a workbook with Aspose.Cells .NET | Export Excel shape text to a CSV file using Aspose.Cells in a console application | Read TextBox shape content from multiple worksheets with Aspose.Cells for .NET | Save extracted textbox values to UTF-8 CSV using Aspose.Cells C#
// Tags: Aspose.Cells TextBoxCollection iteration | export textbox text to CSV C# | read Excel shape text Aspose.Cells | C# write UTF-8 CSV from Excel shapes | console app Aspose.Cells textbox extraction

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook with Aspose.Cells, loops through each worksheet's TextBoxCollection, captures the Text property of every TextBox (using an empty string for null values), and writes the collected strings to a UTF‑8 CSV file with a header row, properly escaping double quotes.
class Program
{
    static void Main(string[] args)
    {
        // Expect two arguments: input Excel file path and output CSV file path
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: Program.exe <input_excel_file> <output_csv_file>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        // Validate input file existence
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The input file \"{inputPath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // List to hold extracted textbox texts
            List<string> extractedTexts = new List<string>();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Access the collection of TextBox shapes on the worksheet
                TextBoxCollection textBoxes = worksheet.TextBoxes;

                // Loop through each TextBox in the collection
                for (int i = 0; i < textBoxes.Count; i++)
                {
                    TextBox tb = textBoxes[i];
                    // The Text property may be null; treat null as empty string
                    string txt = tb.Text ?? string.Empty;
                    extractedTexts.Add(txt);
                }
            }

            // Write the collected texts to a CSV file
            using (StreamWriter writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8))
            {
                // Optional header row
                writer.WriteLine("TextBoxText");

                foreach (string txt in extractedTexts)
                {
                    // Escape double quotes by doubling them, then wrap the field in quotes
                    string escaped = txt.Replace("\"", "\"\"");
                    writer.WriteLine($"\"{escaped}\"");
                }
            }

            Console.WriteLine($"Extraction complete. {extractedTexts.Count} textbox(es) written to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Output full exception details for troubleshooting
            Console.WriteLine("An error occurred during processing:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
