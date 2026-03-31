using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDynamicRows
{
    public class Person
    {
        public string Name { get; set; } = null!;
        public int Age { get; set; }
        public DateTime BirthDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            string inputFile = "input.xlsx";
            Workbook workbook = new Workbook(inputFile);
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, BirthDate = new DateTime(1993, 5, 12) },
                new Person { Name = "Bob",   Age = 45, BirthDate = new DateTime(1978, 11, 3) },
                new Person { Name = "Carol", Age = 27, BirthDate = new DateTime(1996, 2, 20) }
            };

            ImportTableOptions options = new ImportTableOptions
            {
                IsFieldNameShown = true,
                TotalRows = people.Count,
                InsertRows = true,
                DateFormat = "yyyy-MM-dd",
                ConvertNumericData = true,
                CheckMergedCells = true
            };

            int startRow = 0;
            int startColumn = 0;

            int importedRows = cells.ImportCustomObjects(people, startRow, startColumn, options);

            int columnCount = 3; // Name, Age, BirthDate
            int totalRows = importedRows; // ImportCustomObjects already accounts for header when IsFieldNameShown is true

            AsposeRange dataRange = cells.CreateRange(startRow, startColumn, totalRows, columnCount);
            cells.AddRange(dataRange);

            string outputFile = "output.xlsx";
            workbook.Save(outputFile);
        }
    }
}