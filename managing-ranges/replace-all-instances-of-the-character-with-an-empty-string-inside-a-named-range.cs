// Title: Remove asterisk (*) from every cell in a named range with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, defines a named range (A1:A2), scans each cell in the range, strips any '*' characters from string values, and saves the modified file.
// Keywords: Aspose.Cells replace character | remove asterisk named range | C# Aspose.Cells replace string | Excel named range replace | Aspose.Cells .NET remove symbol | replace character in cells | named range string manipulation | Aspose.Cells Range.Replace | Excel asterisk removal
// Common Searches: Aspose.Cells remove * from named range | C# replace character in Excel named range | How to delete asterisk in cells using Aspose.Cells | Iterate named range cells Aspose.Cells | Replace specific symbol in Excel with Aspose.Cells
// Developer Intent: Delete all asterisk characters from cells belonging to a specific named range.
// Use Cases: Clean imported datasets by eliminating stray '*' symbols from a defined range before analysis. | Prepare financial statements by removing placeholder asterisks from a targeted block of cells. | Sanitize user‑entered text in a named range to ensure consistent formatting across the worksheet. | Automate data‑preprocessing for downstream reporting tools that cannot handle special characters.
// AI Prompts: Write C# code using Aspose.Cells to remove a given character from every cell in a named range and save the workbook. | Show how to use the Aspose.Cells Range.Replace method to delete asterisks from a named range in one call. | Explain the steps to retrieve a named range address, iterate its cells, and perform string manipulation with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsReplaceInNamedRange
{
    // This C# example creates a workbook, defines a named range (A1:A2), scans each cell in the range, strips any '*' characters from string values, and saves the modified file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data containing the '*' character
                sheet.Cells["A1"].PutValue("A*B");
                sheet.Cells["A2"].PutValue("*C*");
                sheet.Cells["B1"].PutValue("NoStar"); // outside the named range

                // Define a named range that covers A1:A2
                AsposeRange namedRange = sheet.Cells.CreateRange("A1:A2");
                namedRange.Name = "MyRange";

                // Retrieve the named range definition
                Name rangeName = workbook.Worksheets.Names["MyRange"];
                // RefersTo looks like "=Sheet1!$A$1:$A$2"
                string refersTo = rangeName.RefersTo;

                // Extract the address part (e.g., "$A$1:$A$2")
                string address = refersTo.Substring(refersTo.IndexOf('!') + 1).TrimStart('=');

                // Create a Range object from the address
                AsposeRange range = sheet.Cells.CreateRange(address);

                // Iterate through each cell in the range and remove '*'
                for (int i = 0; i < range.RowCount; i++)
                {
                    for (int j = 0; j < range.ColumnCount; j++)
                    {
                        Cell cell = sheet.Cells[range.FirstRow + i, range.FirstColumn + j];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string original = cell.StringValue;
                            if (original.Contains("*"))
                            {
                                // Replace '*' with an empty string
                                string updated = original.Replace("*", string.Empty);
                                cell.PutValue(updated);
                            }
                        }
                    }
                }

                // Determine output file path
                string outputPath = "ReplacedInNamedRange.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (lifecycle rule)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
