using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace DynamicRowsFromCollection
{
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public DateTime BirthDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, BirthDate = new DateTime(1993, 5, 12) },
                new Person { Name = "Bob",   Age = 45, BirthDate = new DateTime(1978, 11, 3) },
                new Person { Name = "Carol", Age = 27, BirthDate = new DateTime(1996, 2, 20) }
            };

            int lastRowWithData = cells.GetLastDataRow(0);
            int startRow = lastRowWithData + 1;

            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = false,
                InsertRows = true,
                DateFormat = "yyyy-MM-dd",
                ConvertNumericData = true
            };

            worksheet.Cells.ImportCustomObjects((ICollection)people, startRow, 0, importOptions);

            workbook.Save("output.xlsx");
        }
    }
}