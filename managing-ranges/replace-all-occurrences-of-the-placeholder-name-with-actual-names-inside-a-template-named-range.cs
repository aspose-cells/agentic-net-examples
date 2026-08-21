// Title: Replace {{Name}} placeholder in a named range using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, defines a named range "TemplateRange" (A1:A3), and iterates each cell to replace every "{{Name}}" token with a concrete value (e.g., "John Doe"). The updated workbook is then saved as an Excel file.
// Keywords: Aspose.Cells replace placeholder | named range text replacement C# | Aspose.Cells template processing | Excel placeholder substitution .NET | mail merge Aspose.Cells | replace {{Name}} token in Excel | C# Aspose.Cells named range iteration
// Common Searches: How to replace a placeholder in a named range with Aspose.Cells for .NET | Iterate cells of a named range to modify string values in C# | Replace {{Name}} token in Excel template using Aspose.Cells | Aspose.Cells C# replace placeholder and save workbook | Mail‑merge with named ranges in Aspose.Cells
// Developer Intent: Replace all occurrences of the {{Name}} placeholder inside the "TemplateRange" named range with a specific value and save the workbook.
// Use Cases: Generate personalized letters by substituting {{Name}} in a named range before saving the file. | Perform a mail‑merge operation where each recipient's name is inserted into the "TemplateRange" of a template workbook. | Update invoice or receipt templates by replacing placeholder tokens in a defined named range with actual customer data.
// AI Prompts: Provide C# code using Aspose.Cells to replace a custom placeholder in a named range and save the workbook. | Explain how to retrieve a named range from a workbook and iterate its cells to modify string values with Aspose.Cells. | Show how to handle multiple occurrences of the same placeholder within a single cell using Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsReplacePlaceholderInNamedRange
{
    // Creates a workbook, defines a named range "TemplateRange" (A1:A3), and iterates each cell to replace every "{{Name}}" token with a concrete value (e.g., "John Doe"). The updated workbook is then saved as an Excel file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some cells with the placeholder "{{Name}}"
                sheet.Cells["A1"].PutValue("Dear {{Name}},");
                sheet.Cells["A2"].PutValue("Your order has been shipped.");
                sheet.Cells["A3"].PutValue("Thank you, {{Name}}!");

                // Create a named range that covers the cells with the placeholder
                // The range will be named "TemplateRange"
                AsposeRange templateRange = sheet.Cells.CreateRange("A1", "A3");
                templateRange.Name = "TemplateRange";

                // Retrieve the Name object from the workbook's name collection
                Name nameObj = workbook.Worksheets.Names["TemplateRange"];

                // Get the actual Range object associated with the name
                AsposeRange range = nameObj.GetRange();

                // Define the actual name that will replace the placeholder
                string actualName = "John Doe";

                // Iterate through each cell in the range and replace the placeholder
                foreach (Cell cell in range)
                {
                    // Only process cells that contain string data
                    if (cell.Type == CellValueType.IsString)
                    {
                        string cellText = cell.StringValue;
                        if (cellText.Contains("{{Name}}"))
                        {
                            // Replace the placeholder with the actual name
                            string newText = cellText.Replace("{{Name}}", actualName);
                            cell.PutValue(newText);
                        }
                    }
                }

                // Save the workbook to a file
                string outputPath = "TemplateRange_Replaced.xlsx";
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
