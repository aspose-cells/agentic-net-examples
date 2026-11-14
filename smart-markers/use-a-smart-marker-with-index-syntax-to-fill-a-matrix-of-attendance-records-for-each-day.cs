// Title: Create an attendance matrix in Excel using Aspose.Cells smart markers with index syntax in C#
// Common Searches: C# Aspose.Cells smart markers index syntax example for 2D array | How to generate an attendance report in Excel with Aspose.Cells WorkbookDesigner | Using WorkbookDesigner SetDataSource with a custom matrix object in C# | Fill Excel cells from a two‑dimensional string array using Aspose.Cells | Insert smart markers with row and column placeholders in Aspose.Cells
// Tags: Aspose.Cells smart markers index syntax | WorkbookDesigner SetDataSource matrix | C# generate attendance report Excel | populate Excel from 2D string array | Excel smart marker placeholders i j

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, adds person and day headers, places smart markers that reference a 2‑D string array via index syntax (${i}, ${j}), sets the AttendanceMatrix object as the "Matrix" data source with WorkbookDesigner, processes the markers to fill the attendance values, and saves the file as AttendanceReport.xlsx.
public class AttendanceMatrix
{
    // 2‑dimensional array holding attendance status
    public string[,] Values;
    public AttendanceMatrix(string[,] values) { Values = values; }
}

class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Define headers ----------
            string[] persons = { "Alice", "Bob", "Charlie" };
            string[] days    = { "Mon", "Tue", "Wed" };

            // Top‑left corner label
            sheet.Cells[0, 0].PutValue("Person/Day");

            // Day headers (first row)
            for (int j = 0; j < days.Length; j++)
                sheet.Cells[0, j + 1].PutValue(days[j]);

            // Person names (first column)
            for (int i = 0; i < persons.Length; i++)
                sheet.Cells[i + 1, 0].PutValue(persons[i]);

            // ---------- Insert smart markers using index syntax ----------
            // Each cell in the matrix receives a marker that references the 2‑D array
            for (int i = 0; i < persons.Length; i++)
            {
                for (int j = 0; j < days.Length; j++)
                {
                    // ${i} and ${j} are placeholders for the current row/column index
                    sheet.Cells[i + 1, j + 1].PutValue("&=Matrix.Values[${i}][${j}]");
                }
            }

            // ---------- Prepare the data source ----------
            string[,] attendance = new string[persons.Length, days.Length];

            // Sample attendance data
            attendance[0, 0] = "Present";  attendance[0, 1] = "Absent";   attendance[0, 2] = "Present";
            attendance[1, 0] = "Absent";   attendance[1, 1] = "Present";  attendance[1, 2] = "Present";
            attendance[2, 0] = "Present";  attendance[2, 1] = "Present";  attendance[2, 2] = "Absent";

            AttendanceMatrix matrix = new AttendanceMatrix(attendance);

            // ---------- Process smart markers ----------
            // Use WorkbookDesigner (the correct API for smart markers)
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            // Correct overload: first argument is the name used in the smart marker, second is the data source object
            designer.SetDataSource("Matrix", matrix);
            designer.Process();

            // ---------- Save the result ----------
            string outputPath = "AttendanceReport.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Report saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
