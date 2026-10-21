// Title: Replace placeholder text throughout an entire Excel workbook using a regular expression with Aspose.Cells for .NET (C#)
// Common Searches: how to use Aspose.Cells C# to replace a placeholder string in all worksheets | regex find and replace across an entire Excel file with Aspose.Cells | bulk replace {PLACEHOLDER} in every cell using Aspose.Cells .NET | iterate through all cells in a workbook and update string values with Aspose.Cells | Aspose.Cells C# replace text in password‑protected workbook using regex
// Tags: regex replace placeholder in Aspose.Cells workbook | bulk cell text replacement Aspose.Cells .NET | iterate worksheets update string cells Aspose.Cells | search and replace across Excel workbook using Aspose.Cells | replace placeholder across all worksheets C# | Aspose.Cells workbook text substitution with regular expression

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an Excel file (creating a sample workbook if the file is missing), iterates over every worksheet and cell, applies a regular expression to replace all occurrences of the placeholder "{PLACEHOLDER}" with "ActualValue" in string cells, and saves the modified workbook. It also demonstrates handling password‑protected files and load/save errors using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Ensure the input file exists; if not, create a simple workbook for demonstration.
            if (!File.Exists(inputPath))
            {
                workbook = CreateSampleWorkbook();
                workbook.Save(inputPath);
            }
            else
            {
                // Load the workbook; if it is password‑protected, an exception will be caught below.
                workbook = new Workbook(inputPath);
            }
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException for password‑protected files.
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine($"The workbook \"{inputPath}\" is password protected. Unable to load it.");
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Define the regular expression pattern to match the placeholder (e.g., "{PLACEHOLDER}")
        string pattern = @"\{PLACEHOLDER\}";
        string replacement = "ActualValue";
        Regex regex = new Regex(pattern);

        // Iterate through all worksheets and replace placeholder text in string cells
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        string replaced = regex.Replace(original, replacement);
                        if (!original.Equals(replaced))
                        {
                            cell.PutValue(replaced);
                        }
                    }
                }
            }
        }

        try
        {
            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }

    // Creates a simple workbook with a placeholder for demonstration purposes.
    private static Workbook CreateSampleWorkbook()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("This is a {PLACEHOLDER} example.");
        ws.Cells["B2"].PutValue("{PLACEHOLDER} appears again.");
        return wb;
    }
}
