// Title: C# method to populate an Excel template from a dictionary of cell addresses or named ranges and return the workbook as an XLSX MemoryStream using Aspose.Cells
// AI Prompts: Generate a C# function that accepts an IDictionary<string, object> and a path to an .xlsx template, writes each entry to the matching cell address or named range with Aspose.Cells, and returns the filled workbook as a MemoryStream. | Write code that loads an Excel template via Aspose.Cells, iterates over a key‑value collection, updates cells or named ranges accordingly, saves the workbook to a MemoryStream in XLSX format, and includes argument validation and error handling.
// Common Searches: asp.net populate excel template from dictionary using Aspose.Cells | c# fill named ranges in an xlsx template and get MemoryStream | how to write values to specific cells in a template workbook with Aspose.Cells | return populated workbook as stream without saving to disk Aspose.Cells | load excel template and map keys to cell addresses c# aspose.cells
// Tags: populate workbook from dictionary Aspose.Cells | write values to named ranges C# | export workbook to MemoryStream XLSX | load Excel template Aspose.Cells | map cell addresses to values C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The PopulateWorkbook method loads an Excel template, iterates over a IDictionary<string, object> where each key is either a cell address or a named range, writes the corresponding value using Aspose.Cells, saves the result as an XLSX MemoryStream, and provides comprehensive validation and exception handling.
public static class WorkbookHelper
{
    /// <param name="data">
    /// Expected to be a dictionary where the key is either a cell address (e.g., "A1") or a named range defined in the template,
    /// and the value is the content to write into that cell/range.
    /// </param>
    /// <param name="templatePath">Full path to the Excel template file.</param>
    /// <returns>A MemoryStream containing the populated workbook (XLSX format).</returns>
    public static MemoryStream PopulateWorkbook(object data, string templatePath)
    {
        if (string.IsNullOrWhiteSpace(templatePath))
            throw new ArgumentException("Template path must be provided.", nameof(templatePath));

        if (!File.Exists(templatePath))
            throw new FileNotFoundException("The specified template file was not found.", templatePath);

        try
        {
            // Load the template workbook from the specified path.
            Workbook workbook = new Workbook(templatePath);

            // Get the first worksheet (modify if a different sheet is required).
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the data object is a dictionary we can iterate.
            if (data is IDictionary<string, object> dict)
            {
                foreach (KeyValuePair<string, object> entry in dict)
                {
                    string key = entry.Key;
                    object value = entry.Value;

                    // Try to resolve the key as a named range first.
                    AsposeRange namedRange = workbook.Worksheets.GetRangeByName(key);
                    if (namedRange != null)
                    {
                        // Write the value to the first cell of the named range.
                        Cell firstCell = namedRange[0, 0];
                        firstCell.PutValue(value);
                    }
                    else
                    {
                        // Treat the key as a regular cell address (e.g., "B2").
                        Cell cell = worksheet.Cells[key];
                        cell.PutValue(value);
                    }
                }
            }
            else
            {
                throw new ArgumentException("Data must be a dictionary with string keys and object values.", nameof(data));
            }

            // Save the populated workbook into a memory stream.
            MemoryStream stream = new MemoryStream();
            workbook.Save(stream, SaveFormat.Xlsx);
            stream.Position = 0; // Reset stream position for reading.

            return stream;
        }
        catch (Exception ex)
        {
            // Wrap any exception to provide context while preserving the original stack trace.
            throw new InvalidOperationException("Failed to populate the workbook.", ex);
        }
    }
}

// Minimal entry point to satisfy the compiler.
public class Program
{
    public static void Main()
    {
        // Example usage (optional):
        // var data = new Dictionary<string, object>
        // {
        //     { "A1", "Sample Text" },
        //     { "MyNamedRange", 12345 }
        // };
        // string templatePath = "template.xlsx";
        // using (MemoryStream ms = WorkbookHelper.PopulateWorkbook(data, templatePath))
        // {
        //     // Process the resulting stream as needed.
        // }
    }
}
