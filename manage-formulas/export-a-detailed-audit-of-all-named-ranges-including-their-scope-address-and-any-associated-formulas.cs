// Title: Generate a CSV report of all named ranges in an Excel workbook, including scope, address, and formulas, using Aspose.Cells for .NET
// Common Searches: C# code to export all named ranges with their scope and formulas to a CSV file using Aspose.Cells | how to list workbook‑wide and worksheet‑specific named ranges with Aspose.Cells .NET | retrieve RefersTo address and formula of defined names in an Excel file via Aspose.Cells | Aspose.Cells example for auditing named ranges and saving results as CSV | C# iterate through workbook.Worksheets.Names and get scope using Aspose.Cells
// Tags: export named ranges to CSV Aspose.Cells | named range scope detection Aspose.Cells | retrieve defined name formulas Aspose.Cells | C# Aspose.Cells list workbook names | Excel named range audit Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Text;
using System.Reflection;

// The sample loads an Excel workbook with Aspose.Cells, iterates over workbook.Worksheets.Names, determines each name's scope (workbook or specific worksheet) via reflection, captures the RefersTo address, extracts a formula when the address starts with '=', and writes a UTF‑8 CSV file containing columns for Name, Scope, Address, and Formula.
class NamedRangesAudit
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "NamedRangesAudit.csv";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Build CSV content
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Name,Scope,Address,Formula");

            // Iterate through all defined names in the workbook
            foreach (Name name in workbook.Worksheets.Names)
            {
                // Determine the scope: workbook-wide or specific worksheet
                string scope = "Workbook";
                try
                {
                    // Attempt to read the Worksheet property via reflection (covers older API versions)
                    PropertyInfo wsProp = name.GetType().GetProperty("Worksheet");
                    if (wsProp != null)
                    {
                        Worksheet ws = wsProp.GetValue(name) as Worksheet;
                        if (ws != null)
                            scope = ws.Name;
                    }
                }
                catch { /* ignore reflection errors */ }

                // The address or formula the name refers to
                string address = name.RefersTo ?? string.Empty;

                // If the address starts with '=', treat it as a formula
                string formula = (!string.IsNullOrEmpty(address) && address.StartsWith("=")) ? address : string.Empty;

                // Append CSV line with proper escaping
                sb.AppendLine($"{Escape(name.Text)},{Escape(scope)},{Escape(address)},{Escape(formula)}");
            }

            // Write the CSV audit to a file
            File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);
            Console.WriteLine($"Audit saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to escape CSV fields containing commas, quotes, or newlines
    static string Escape(string field)
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
