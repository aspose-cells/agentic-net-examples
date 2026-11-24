// Title: Generate a CSV report of all named ranges with their scope, address, and formula using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export named ranges to CSV with scope and formula details | How to list workbook‑wide and worksheet‑specific named ranges using Aspose.Cells .NET | Retrieve formula expressions from defined names in an Excel file with Aspose.Cells | Iterate through Worksheets.Names collection and save results to CSV in C# | Export named range definitions including RefersTo to a CSV file with Aspose.Cells
// Tags: named range csv export aspose.cells | aspose.cells retrieve defined name scope | aspose.cells extract formula from named range | c# audit excel named ranges aspose.cells | workbook names enumeration aspose.cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates over the Worksheets.Names collection, determines each defined name's scope (global workbook or specific worksheet), parses the RefersTo string to separate plain cell addresses from formula expressions, escapes CSV fields, and writes a UTF‑8 CSV file containing columns Name, Scope, Address, and Formula.
class NamedRangesAudit
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "NamedRangesAudit.csv";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare a StringBuilder for CSV output
            StringBuilder csv = new StringBuilder();

            // Write CSV header
            csv.AppendLine("Name,Scope,Address,Formula");

            // Iterate through all defined names (named ranges)
            foreach (var definedNameObj in workbook.Worksheets.Names)
            {
                // Use dynamic to safely access members that may differ between library versions
                dynamic definedName = definedNameObj;

                // Determine the scope: Workbook (global) or specific worksheet name
                string scope = "Workbook";
                try
                {
                    // If the library exposes IsGlobal and Worksheet, use them
                    bool isGlobal = definedName.IsGlobal;
                    if (!isGlobal && definedName.Worksheet != null)
                    {
                        scope = definedName.Worksheet.Name ?? "Workbook";
                    }
                }
                catch
                {
                    // Fallback: keep default "Workbook" scope
                }

                // Get the defined name text
                string name = string.Empty;
                try { name = definedName.Name ?? string.Empty; } catch { }

                // The RefersTo property contains the definition, e.g. =Sheet1!$A$1:$B$2 or =SUM(Sheet1!$A$1:$A$10)
                string refersTo = string.Empty;
                try { refersTo = definedName.RefersTo ?? string.Empty; } catch { }

                // Remove leading '=' for address extraction
                string address = refersTo.StartsWith("=") ? refersTo.Substring(1) : refersTo;

                // Identify if the definition is a formula (contains '(') or a plain range
                string formula = address.Contains("(") ? address : string.Empty;

                // If it's a plain range, keep the address; otherwise leave address empty
                if (!string.IsNullOrEmpty(formula))
                    address = string.Empty;

                // Escape commas and quotes in fields if necessary
                string nameEscaped = EscapeCsv(name);
                string scopeEscaped = EscapeCsv(scope);
                string addressEscaped = EscapeCsv(address);
                string formulaEscaped = EscapeCsv(formula);

                // Append the CSV line
                csv.AppendLine($"{nameEscaped},{scopeEscaped},{addressEscaped},{formulaEscaped}");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Export the audit to a CSV file
            File.WriteAllText(outputPath, csv.ToString(), Encoding.UTF8);
            Console.WriteLine($"Audit completed. CSV saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to escape CSV fields containing commas, quotes, or newlines
    private static string EscapeCsv(string field)
    {
        if (string.IsNullOrEmpty(field))
            return string.Empty;

        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            field = field.Replace("\"", "\"\"");
            return $"\"{field}\"";
        }
        return field;
    }
}
