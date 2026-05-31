using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

class CellsEnumeratorToJson
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Age");
        cells["A2"].PutValue("John");
        cells["B2"].PutValue(30);
        cells["A3"].PutValue("Alice");
        cells["B3"].PutValue(25);

        // Obtain the enumerator for the Cells collection
        IEnumerator enumerator = cells.GetEnumerator();

        // Collect JSON representation of each cell that has a value
        List<string> cellJsonList = new List<string>();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            if (cell.Value != null)
            {
                // Cell.ToJson returns a JSON object for the individual cell
                cellJsonList.Add(cell.ToJson());
            }
        }

        // Combine individual cell JSON objects into a JSON array
        string jsonResult = "[" + string.Join(",", cellJsonList) + "]";

        // Output the final JSON string
        Console.WriteLine(jsonResult);
    }
}