using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class JsonMultipleWorksheetsDemo
    {
        public static void Run()
        {
            // Path to the JSON file that contains an array of tables.
            string jsonFilePath = "input.json";

            // Configure JSON load options:
            // - MultipleWorksheets = true : each top‑level array element will be placed in its own worksheet.
            JsonLoadOptions loadOptions = new JsonLoadOptions
            {
                MultipleWorksheets = true
            };

            // Load the JSON data into a workbook using the specified options.
            Workbook workbook = new Workbook(jsonFilePath, loadOptions);

            // Optional: verify the number of worksheets created.
            Console.WriteLine($"Worksheets created: {workbook.Worksheets.Count}");

            // Save the workbook to an Excel file.
            workbook.Save("output.xlsx");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            JsonMultipleWorksheetsDemo.Run();
        }
    }
}