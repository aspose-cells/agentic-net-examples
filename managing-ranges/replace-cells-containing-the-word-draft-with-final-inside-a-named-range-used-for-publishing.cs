// Title: Replace "Draft" with "Final" in a named range using Aspose.Cells for .NET (C#)
// Description: Loads an Excel template, locates the named range "PublishRange", parses its RefersTo address, creates a Range object, swaps every occurrence of the word "Draft" with "Final" in string cells, and saves the result as "Published.xlsx".
// Keywords: Aspose.Cells replace text named range | C# update Excel cells Aspose | replace Draft Final Aspose.Cells | named range string replacement .NET | Excel workbook text replace Aspose
// Common Searches: Aspose.Cells replace word in named range C# | How to change text in a specific Excel named range using .NET | C# replace Draft with Final in Excel range Aspose | Programmatically edit cells inside a named range Aspose.Cells
// Developer Intent: Swap every "Draft" occurrence for "Final" within the cells of a defined named range and persist the changes to a new workbook.
// Use Cases: Prepare a publish‑ready copy of a template by converting draft markers to final text in a designated range. | Automate the finalization step of a report section that is isolated by a named range before distribution. | Batch‑process status flags or version labels across a predefined area of a workbook as part of a data pipeline.
// AI Prompts: Write C# code with Aspose.Cells that finds a named range and replaces a substring in all string cells. | Explain how to extract the worksheet name and address from a named range's RefersTo property in Aspose.Cells. | Suggest robust error‑handling patterns when modifying cell values inside a named range with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace ReplaceDraftInNamedRange
{
    // Loads an Excel template, locates the named range "PublishRange", parses its RefersTo address, creates a Range object, swaps every occurrence of the word "Draft" with "Final" in string cells, and saves the result as "Published.xlsx".
    class Program
    {
        static void Main()
        {
            const string templatePath = "Template.xlsx";
            const string outputPath = "Published.xlsx";
            const string rangeName = "PublishRange";

            // Verify template file exists
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file '{templatePath}' not found.");
                return;
            }

            try
            {
                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(templatePath);

                // Retrieve the named range definition
                Name namedRange = workbook.Worksheets.Names[rangeName];
                if (namedRange == null)
                {
                    Console.WriteLine($"Named range '{rangeName}' not found.");
                    return;
                }

                // The RefersTo string looks like "=Sheet1!$A$1:$C$10"
                string refersTo = namedRange.RefersTo.TrimStart('=');
                int exclPos = refersTo.IndexOf('!');
                if (exclPos < 0)
                {
                    Console.WriteLine("Invalid RefersTo format.");
                    return;
                }

                string sheetName = refersTo.Substring(0, exclPos);
                string address = refersTo.Substring(exclPos + 1);

                // Get the worksheet that contains the range
                Worksheet ws = workbook.Worksheets[sheetName];
                if (ws == null)
                {
                    Console.WriteLine($"Worksheet '{sheetName}' not found.");
                    return;
                }

                // Create the Range object from the address (use fully qualified type to avoid ambiguity)
                Aspose.Cells.Range range = ws.Cells.CreateRange(address);

                // Iterate through each cell in the range and replace "Draft" with "Final"
                foreach (Cell cell in range)
                {
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue;
                        if (text.Contains("Draft"))
                        {
                            string newText = text.Replace("Draft", "Final");
                            cell.PutValue(newText);
                        }
                    }
                }

                // Save the modified workbook (lifecycle rule: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
