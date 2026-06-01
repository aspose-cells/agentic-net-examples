using System;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing XLS workbook
            string sourcePath = "input.xls";

            // Load the workbook using the string constructor (loads the file)
            Workbook workbook = new Workbook(sourcePath);

            // Access the DataMashup property
            DataMashup dataMashup = workbook.DataMashup;

            // Demonstrate usage of the DataMashup object
            if (dataMashup != null)
            {
                // Retrieve all Power Query formulas contained in the workbook
                var powerQueryFormulas = dataMashup.PowerQueryFormulas;

                // Output the number of Power Query formulas found
                Console.WriteLine($"Number of Power Query formulas: {powerQueryFormulas.Count}");

                // Optionally, list the names of the formulas
                foreach (var formula in powerQueryFormulas)
                {
                    Console.WriteLine($"Formula Name: {formula.Name}");
                }
            }
            else
            {
                Console.WriteLine("DataMashup property is null.");
            }

            // Save the workbook to a new file (optional)
            workbook.Save("output.xls");
        }
    }
}