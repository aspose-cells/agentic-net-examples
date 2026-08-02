using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the JSON file that contains an array of tables
        string jsonFilePath = "input.json";

        // Configure JSON load options:
        // - MultipleWorksheets = true creates a separate worksheet for each top‑level array element
        // - LayoutOptions.ArrayAsTable = true treats each array as a table (rows/columns)
        // - LayoutOptions.IgnoreTitle = true skips title rows if present
        JsonLoadOptions loadOptions = new JsonLoadOptions
        {
            MultipleWorksheets = true,
            LayoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,
                IgnoreTitle = true
            }
        };

        // Load the JSON data into a workbook using the specified options
        Workbook workbook = new Workbook(jsonFilePath, loadOptions);

        // Save the resulting workbook to an Excel file
        workbook.Save("output.xlsx");
    }
}