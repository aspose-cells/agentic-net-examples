// Title: Load an Excel template workbook from a file or stream with Aspose.Cells in C# before using smart markers
// Common Searches: how to load an existing Excel template into an Aspose.Cells workbook in C# | Aspose.Cells load workbook from stream when template file may be missing | C# create a new workbook with Aspose.Cells if Template.xlsx does not exist | load Excel template and fallback to a new workbook before applying smart markers using Aspose.Cells | read Excel template from file path with Aspose.Cells C# example
// Tags: load workbook from file Aspose.Cells C# | load workbook from stream Aspose.Cells C# | fallback create new workbook Aspose.Cells | initialize workbook for smart markers Aspose.Cells | check template existence Aspose.Cells C# | conditional workbook loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks whether "Template.xlsx" exists, loads it into an Aspose.Cells Workbook (or loads from a stream), and if the file is absent creates a new workbook with a default worksheet, writes sample data, and saves the result as "Result.xlsx" ready for smart marker processing.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the Excel template.
            string templatePath = "Template.xlsx";

            // Load the workbook from the template if it exists; otherwise create a new workbook.
            Workbook workbook;
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                Console.WriteLine($"Template file \"{templatePath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
                // Add a default worksheet if none exists.
                if (workbook.Worksheets.Count == 0)
                {
                    workbook.Worksheets.Add("Sheet1");
                }
            }

            // -------------------------------------------------
            // Example data source (replace with your own logic)
            // -------------------------------------------------
            var dataSource = new
            {
                Name = "John Doe",
                Age = 30,
                Department = "Sales"
            };

            // -------------------------------------------------
            // Populate the workbook with the data source.
            // (Smart Markers are not used here to avoid missing API.)
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Write headers.
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["C1"].PutValue("Department");

            // Write data.
            cells["A2"].PutValue(dataSource.Name);
            cells["B2"].PutValue(dataSource.Age);
            cells["C2"].PutValue(dataSource.Department);

            // Save the processed workbook.
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to \"{resultPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
