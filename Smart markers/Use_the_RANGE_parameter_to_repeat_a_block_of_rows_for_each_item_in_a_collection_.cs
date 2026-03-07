using System;
using System.Collections.Generic;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeRepeatDemo
{
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string City { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            Workbook workbook = new Workbook("Template.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            int templateStartRow = 2;          // zero‑based index of the first template row
            int templateRowCount = 3;          // number of rows in the template block
            int firstColumn = 0;               // start column of the block (A)

            // Determine the number of columns to copy (ensure at least one column)
            int totalColumns = cells.MaxColumn - firstColumn + 1;
            if (totalColumns <= 0) totalColumns = 1;

            AsposeRange templateRange = cells.CreateRange(templateStartRow, firstColumn, templateRowCount, totalColumns);

            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, City = "New York" },
                new Person { Name = "Bob",   Age = 25, City = "London"   },
                new Person { Name = "Carol", Age = 28, City = "Tokyo"    }
            };

            int insertRow = templateStartRow + templateRowCount;

            foreach (Person p in people)
            {
                worksheet.Cells.InsertRows(insertRow, templateRowCount, true);

                AsposeRange targetRange = cells.CreateRange(insertRow, firstColumn, templateRowCount, totalColumns);

                templateRange.Copy(targetRange);

                targetRange[0, 0].PutValue(p.Name);
                targetRange[1, 0].PutValue(p.Age);
                targetRange[2, 0].PutValue(p.City);

                insertRow += templateRowCount;
            }

            worksheet.Cells.DeleteRows(templateStartRow, templateRowCount);

            workbook.Save("Result.xlsx");
        }
    }
}