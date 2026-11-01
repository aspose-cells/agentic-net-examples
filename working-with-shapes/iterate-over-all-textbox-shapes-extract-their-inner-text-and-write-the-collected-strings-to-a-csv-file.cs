// Title: Extract text from all TextBox shapes in an Excel workbook and write it to a CSV file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells read TextBox shape text from each worksheet | How to export Excel TextBox contents to a CSV file with Aspose.Cells | Iterate through shapes in a workbook and collect textbox values using Aspose.Cells .NET | Save extracted TextBox strings from a password‑protected Excel file to CSV in C# | Aspose.Cells extract shape text and generate CSV output programmatically
// Tags: Aspose.Cells extract textbox shape text | Aspose.Cells write shape text to CSV | iterate worksheet shapes Aspose.Cells | load password protected workbook Aspose.Cells | C# generate CSV from extracted strings

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook (prompting for a password when needed), iterates over every worksheet and its shapes, gathers the Text property of each TextBox shape, ensures the target directory exists, and writes each collected string as a quoted, escaped line to a CSV file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            Workbook workbook = null;

            // Attempt to load the workbook; if password is required, prompt the user
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception)
            {
                Console.Write("The workbook is password protected. Enter password: ");
                string password = Console.ReadLine() ?? string.Empty;

                LoadOptions loadOptions = new LoadOptions { Password = password };
                workbook = new Workbook(inputPath, loadOptions);
            }

            // Collect texts from all TextBox shapes
            List<string> textboxTexts = new List<string>();
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape is TextBox textbox)
                    {
                        textboxTexts.Add(textbox.Text);
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Write the collected strings to a CSV file (one entry per line)
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                foreach (string line in textboxTexts)
                {
                    // Escape double quotes and wrap the field in quotes
                    string escaped = $"\"{line.Replace("\"", "\"\"")}\"";
                    writer.WriteLine(escaped);
                }
            }

            Console.WriteLine($"Extraction complete. Results saved to \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
