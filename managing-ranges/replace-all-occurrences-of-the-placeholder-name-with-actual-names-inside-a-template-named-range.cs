// Title: Replace {{Name}} placeholder with defined name values in all named ranges using Aspose.Cells for .NET
// Description: Loads a template workbook, iterates through each defined name, accesses its range, substitutes any cell containing the "{{Name}}" token with the name's actual text, and saves the updated file. Includes file‑existence verification and safe handling of names that are not linked to a range.
// Keywords: Aspose.Cells | C# | replace placeholder | named range | defined name | Excel template | token replacement | workbook manipulation | .NET | range iteration
// Common Searches: Aspose.Cells replace placeholder in named range | C# replace {{Name}} token in Excel template | How to update defined name values in cells with Aspose.Cells | Iterate named ranges and modify cell text .NET | Replace tokens in Excel named ranges programmatically
// Developer Intent: Replace every "{{Name}}" token inside each named range with the corresponding defined name's text.
// Use Cases: Personalized employee reports – inject employee name into named‑range placeholders. | Client‑specific invoice generation – swap {{Name}} with the client’s name in a template. | Mail‑merge style document creation – automatically fill named‑range placeholders with record values. | Dynamic dashboard generation – replace placeholder with chart titles stored as defined names.
// AI Prompts: Generate C# Aspose.Cells code that scans all defined names, gets their ranges, and replaces "{{Name}}" with the name's text. | Show how to safely skip defined names that do not refer to a range while performing placeholder replacement. | Explain error handling for missing template files when using Aspose.Cells to replace tokens. | Provide a concise snippet that updates cells containing a custom token within named ranges.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsReplacePlaceholderInNamedRanges
{
    // Loads a template workbook, iterates through each defined name, accesses its range, substitutes any cell containing the "{{Name}}" token with the name's actual text, and saves the updated file. Includes file‑existence verification and safe handling of names that are not linked to a range.
    public class Program
    {
        public static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "Result.xlsx";

                // Verify that the template file exists to avoid FileNotFoundException
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Error: The file \"{templatePath}\" was not found.");
                    return;
                }

                // Load the workbook that contains the template named range(s)
                Workbook workbook = new Workbook(templatePath);

                // Get the collection of defined names (named ranges) in the workbook
                NameCollection names = workbook.Worksheets.Names;

                // Iterate through each defined name
                foreach (Name name in names)
                {
                    // Obtain the actual range that the name refers to
                    Aspose.Cells.Range range = name.GetRange();

                    // If the name does not refer to a range, skip it
                    if (range == null) continue;

                    // Iterate through each cell in the range
                    foreach (Cell cell in range)
                    {
                        // Check if the cell contains the placeholder "{{Name}}"
                        string cellText = cell.StringValue;
                        if (!string.IsNullOrEmpty(cellText) && cellText.Contains("{{Name}}"))
                        {
                            // Replace the placeholder with the actual name text
                            string replacedText = cellText.Replace("{{Name}}", name.Text);
                            cell.PutValue(replacedText);
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
