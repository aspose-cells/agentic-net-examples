// Title: C# – Mask “confidential” text in a named range with Aspose.Cells
// Description: Loads an Excel workbook, retrieves the named range "ConfidentialRange", creates a Range object, scans each cell for the word "confidential" (case‑insensitive, partial match) and replaces it with "*****", then saves the updated file.
// Keywords: Aspose.Cells C# | mask confidential text | named range replace | case insensitive replace Excel | redact sensitive data Aspose | search cells in named range | Excel data privacy .NET | replace text in workbook range
// Common Searches: replace word in named range Aspose.Cells | mask confidential data in Excel C# | case insensitive text replace in workbook | iterate cells of a named range Aspose | redact sensitive keywords in Excel using .NET
// Developer Intent: Find cells containing "confidential" inside the named range "ConfidentialRange" and replace the word with a mask.
// Use Cases: Sanitize reports before sharing with external partners. | Automate compliance by redacting confidential keywords in predefined worksheet sections. | Generate clean data extracts for analytics while protecting sensitive terms.
// AI Prompts: Write C# code using Aspose.Cells that loads an Excel file, locates a named range, and replaces every occurrence of a given keyword with asterisks, ignoring case. | Create a method that checks for a named range, builds an Aspose.Cells.Range object, iterates its cells, and masks a specified word with "*****".

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, retrieves the named range "ConfidentialRange", creates a Range object, scans each cell for the word "confidential" (case‑insensitive, partial match) and replaces it with "*****", then saves the updated file.
    public class MaskConfidentialInNamedRange
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "ConfidentialRange"
            Name namedRange = workbook.Worksheets.Names["ConfidentialRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'ConfidentialRange' not found.");
                return;
            }

            // Extract the address without the leading '='
            string rangeAddress = namedRange.RefersTo;
            if (rangeAddress.StartsWith("="))
                rangeAddress = rangeAddress.Substring(1);

            // Assume the range is on the first worksheet (adjust if necessary)
            Worksheet sheet = workbook.Worksheets[0];

            // Create an Aspose.Cells.Range from the address
            Aspose.Cells.Range range = sheet.Cells.CreateRange(rangeAddress);

            // Configure replace options: case‑insensitive, partial match
            ReplaceOptions replaceOptions = new ReplaceOptions
            {
                CaseSensitive = false,
                MatchEntireCellContents = false
            };

            // Iterate through cells in the range and mask the word "confidential"
            foreach (Cell cell in range)
            {
                if (cell.Type == CellValueType.IsString &&
                    cell.StringValue.IndexOf("confidential", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    cell.Replace("confidential", "*****", replaceOptions);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
