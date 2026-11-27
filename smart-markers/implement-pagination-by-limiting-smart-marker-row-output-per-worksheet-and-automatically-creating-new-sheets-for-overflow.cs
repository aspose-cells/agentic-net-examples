// Title: C# pagination of smart marker rows with Aspose.Cells: limit rows per sheet and auto‑create new worksheets
// Common Searches: Aspose.Cells C# paginate smart marker rows by limiting rows per sheet | How to split smart marker data source into multiple worksheets with WorkbookDesigner | Automatically create new Excel sheets when smart marker output exceeds row limit in .NET | C# example for smart marker pagination using template sheet copying | Limit rows per worksheet for smart markers and generate additional pages with Aspose.Cells
// Tags: smart marker pagination Aspose.Cells | row limit per worksheet Aspose.Cells | auto generate worksheets Aspose.Cells | WorkbookDesigner process specific sheet | split list into chunks C# Excel export

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerPaginationDemo
{
    // Simple data model used for the smart markers
    // The sample loads a template workbook containing smart markers, divides a large Person list into chunks that respect a maximum row count per sheet, copies the template sheet for each chunk, binds each chunk to the "Persons" data source, processes only the target worksheet with WorkbookDesigner, and saves the resulting paginated Excel file.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }

        public Person(string name, int age)
        {
            Name = name;
            Age = age;
        }
    }

    class Program
    {
        // Maximum number of data rows that should appear on a single worksheet.
        // Adjust this value according to your pagination requirements.
        private const int MaxRowsPerSheet = 30;

        static void Main(string[] args)
        {
            try
            {
                // Paths to the template and the resulting workbook.
                // The template must contain the smart markers that reference the data source name "Persons".
                // Example of a smart marker row in the template:
                //   A2: &Persons.Name
                //   B2: &Persons.Age
                string templatePath = "Template.xlsx";
                string outputPath = "PaginatedOutput.xlsx";

                // Load the template workbook.
                Workbook workbook = new Workbook(templatePath);

                // Index of the worksheet that holds the smart markers in the template.
                // In most cases this will be the first worksheet (index 0).
                int templateSheetIndex = 0;

                // Generate a sample data source with more rows than MaxRowsPerSheet to demonstrate pagination.
                List<Person> allPersons = GenerateSampleData(95); // 95 rows > 3 pages when MaxRowsPerSheet = 30

                // Split the data source into chunks that fit within the row limit.
                List<List<Person>> dataChunks = SplitList(allPersons, MaxRowsPerSheet);

                // Process each chunk on its own worksheet.
                for (int i = 0; i < dataChunks.Count; i++)
                {
                    int targetSheetIndex;

                    if (i == 0)
                    {
                        // First chunk uses the original template sheet.
                        targetSheetIndex = templateSheetIndex;
                        workbook.Worksheets[targetSheetIndex].Name = $"Page {i + 1}";
                    }
                    else
                    {
                        // Subsequent chunks get a copy of the template sheet.
                        targetSheetIndex = workbook.Worksheets.AddCopy(templateSheetIndex);
                        workbook.Worksheets[targetSheetIndex].Name = $"Page {i + 1}";
                    }

                    // Create a designer for the workbook and bind the current chunk.
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);
                    designer.SetDataSource("Persons", dataChunks[i]);

                    // Process only the target worksheet (preserve unrecognized markers = true).
                    designer.Process(targetSheetIndex, true);
                }

                // Save the paginated workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Paginated workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="count">Number of Person objects to generate.</param>
        /// <returns>List of Person instances.</returns>
        private static List<Person> GenerateSampleData(int count)
        {
            var list = new List<Person>();
            for (int i = 1; i <= count; i++)
            {
                list.Add(new Person($"Person {i}", 20 + (i % 30)));
            }
            return list;
        }

        /// <summary>
        /// Splits a list into smaller sub‑lists each containing at most <paramref name="size"/> elements.
        /// </summary>
        /// <typeparam name="T">Type of the list elements.</typeparam>
        /// <param name="source">Original list to split.</param>
        /// <param name="size">Maximum size of each chunk.</param>
        /// <returns>List of chunks.</returns>
        private static List<List<T>> SplitList<T>(List<T> source, int size)
        {
            var result = new List<List<T>>();
            for (int i = 0; i < source.Count; i += size)
            {
                int chunkSize = Math.Min(size, source.Count - i);
                result.Add(source.GetRange(i, chunkSize));
            }
            return result;
        }
    }
}
