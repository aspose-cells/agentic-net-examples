using System;
using Aspose.Cells;

class QueryTableWorksheetLister
{
    static void Main(string[] args)
    {
        // Ensure a file path is provided
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: QueryTableWorksheetLister <input_excel_file>");
            return;
        }

        string inputPath = args[0];

        // Load the workbook (load rule)
        Workbook workbook = new Workbook(inputPath);

        // Iterate over all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // If the worksheet has one or more query tables, output its name
            if (sheet.QueryTables.Count > 0)
            {
                Console.WriteLine(sheet.Name);
            }
        }

        // Save the workbook (save rule) – no modifications are made, but the rule is applied
        workbook.Save("output.xlsx");
    }
}