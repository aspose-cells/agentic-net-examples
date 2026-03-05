using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsListObjectDemo
{
    // Sample data class
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a list of custom objects
            List<Person> people = new List<Person>
            {
                new Person { Name = "John", Age = 30, City = "New York" },
                new Person { Name = "Alice", Age = 25, City = "London" },
                new Person { Name = "Bob", Age = 28, City = "Sydney" }
            };

            // Import the custom objects into the worksheet.
            // The header row (property names) will be added.
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,   // show property names as header
                InsertRows = true,         // insert rows if needed
                ConvertNumericData = true, // convert numeric strings to numbers
                DateFormat = "yyyy-MM-dd"
            };

            // Import starting at cell A1 (row 0, column 0)
            sheet.Cells.ImportCustomObjects(people, 0, 0, importOptions);

            // Determine the range of the imported data
            int startRow = 0;                     // header row
            int startColumn = 0;                  // column A
            int endRow = people.Count;            // header + data rows (0-based index)
            int endColumn = 2;                    // three columns: Name, Age, City (0,1,2)

            // Add a ListObject (table) over the imported range
            int tableIndex = sheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "PeopleTable";
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Save the workbook as XLSX
            workbook.Save("PeopleListObject.xlsx", SaveFormat.Xlsx);
        }
    }
}