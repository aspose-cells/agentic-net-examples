using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsReplacePlaceholderInNamedRange
{
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

                // Populate a range with the placeholder "{{Name}}"
                sheet.Cells["A1"].PutValue("{{Name}}");
                sheet.Cells["A2"].PutValue("{{Name}}");
                sheet.Cells["A3"].PutValue("{{Name}}");

                // Define a named range that refers to the template cells
                int nameIndex = workbook.Worksheets.Names.Add("TemplateRange");
                // The RefersTo formula must start with '='
                workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$A$3";

                // Array of actual names that will replace the placeholder
                string[] actualNames = { "Alice", "Bob", "Charlie" };

                // Retrieve the named range object
                Name templateName = workbook.Worksheets.Names["TemplateRange"];
                AsposeRange templateRange = templateName.GetRange();

                // Iterate through each cell in the named range and replace the placeholder
                for (int i = 0; i < templateRange.RowCount; i++)
                {
                    // Since the range is a single column, column index is 0
                    Cell cell = templateRange[i, 0];

                    // Check if the cell contains the placeholder
                    if (cell.StringValue.Contains("{{Name}}"))
                    {
                        // Replace with the corresponding actual name (if available)
                        string newValue = i < actualNames.Length ? actualNames[i] : "Unknown";
                        cell.PutValue(newValue);
                    }
                }

                // Save the workbook to a file
                string outputPath = "TemplateRangeReplaced.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}