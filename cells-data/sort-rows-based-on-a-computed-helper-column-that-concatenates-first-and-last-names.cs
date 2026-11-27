// Title: How to sort Excel rows by a concatenated full name helper column using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# sort worksheet rows by a calculated full name column | use DataSorter to order Excel data on a helper column that concatenates first and last names | remove temporary helper column after sorting with Aspose.Cells .NET | calculate formulas before sorting rows in an Aspose.Cells workbook | define CellArea range for sorting with Aspose.Cells DataSorter
// Tags: DataSorter sort by computed column Aspose.Cells | helper column concatenation first name last name C# | calculate formulas before DataSorter Aspose.Cells | delete temporary column after sorting Aspose.Cells | CellArea range definition for sorting Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSortByHelperColumn
{
    // The example creates a workbook, adds first name, last name, and age columns, builds a helper column that concatenates first and last names with a formula, calculates the formulas, defines a CellArea covering the data, configures DataSorter to sort ascending by the helper column, performs the sort, deletes the temporary helper column, and saves the sorted file as SortedByFullName.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data: First Name (Column A), Last Name (Column B), Age (Column C)
                // Header row
                sheet.Cells["A1"].PutValue("FirstName");
                sheet.Cells["B1"].PutValue("LastName");
                sheet.Cells["C1"].PutValue("Age");

                // Data rows
                sheet.Cells["A2"].PutValue("John");
                sheet.Cells["B2"].PutValue("Doe");
                sheet.Cells["C2"].PutValue(28);

                sheet.Cells["A3"].PutValue("Alice");
                sheet.Cells["B3"].PutValue("Smith");
                sheet.Cells["C3"].PutValue(34);

                sheet.Cells["A4"].PutValue("Bob");
                sheet.Cells["B4"].PutValue("Brown");
                sheet.Cells["C4"].PutValue(22);

                sheet.Cells["A5"].PutValue("Charlie");
                sheet.Cells["B5"].PutValue("Adams");
                sheet.Cells["C5"].PutValue(30);

                // Add a helper column (Column D) that concatenates FirstName and LastName
                // Header for helper column
                sheet.Cells["D1"].PutValue("FullNameHelper");

                // Populate helper column with formulas
                // Rows start at index 1 (second row) because row 0 is header
                for (int row = 1; row <= 4; row++)
                {
                    // Formula: =A{row+1}&B{row+1}
                    // Using CONCATENATE function for clarity (optional)
                    sheet.Cells[row, 3].Formula = $"=A{row + 1}&B{row + 1}";
                }

                // Calculate the formulas so the helper column contains actual values
                workbook.CalculateFormula();

                // Configure the DataSorter to sort by the helper column (column index 3)
                DataSorter sorter = workbook.DataSorter;
                sorter.Order1 = SortOrder.Ascending;   // Ascending order
                sorter.Key1 = 3;                       // Zero‑based column index for column D

                // Define the sort range (including the helper column)
                CellArea sortArea = new CellArea
                {
                    StartRow = 0,          // Include header row
                    EndRow = 4,            // Last data row (row index 4)
                    StartColumn = 0,       // Column A
                    EndColumn = 3          // Column D (helper column)
                };

                // Perform the sort
                sorter.Sort(sheet.Cells, sortArea);

                // Optional: Remove the helper column after sorting
                sheet.Cells.DeleteColumn(3); // Deletes column D (zero‑based index)

                // Save the sorted workbook
                string outputPath = "SortedByFullName.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook sorted successfully and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
