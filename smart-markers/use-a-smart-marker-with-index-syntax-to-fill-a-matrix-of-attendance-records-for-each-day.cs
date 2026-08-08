// Title: C# – Build an Excel Attendance Matrix with Aspose.Cells Smart Markers and Index Syntax
// Description: Demonstrates how to generate an attendance spreadsheet by creating a workbook, adding a header row, inserting smart markers with the [0] index to repeat rows for each DataTable record, naming the marker range "_CellsSmartMarkers", binding the DataTable to a WorkbookDesigner, processing the named range, and saving the result as an XLSX file.
// Keywords: Aspose.Cells | C# | smart markers | index syntax | attendance matrix | DataTable | WorkbookDesigner | named range | Excel export | repeat rows
// Common Searches: Aspose.Cells smart markers repeat rows C# | attendance matrix Excel Aspose.Cells | WorkbookDesigner process named range | C# fill Excel from DataTable using smart markers | index [0] syntax Aspose.Cells
// Developer Intent: Create an Excel attendance report by applying smart markers with the [0] index to auto‑repeat rows from a DataTable.
// Use Cases: Generate daily attendance logs for multiple employees with one row per date. | Export any DataTable (e.g., sales, inventory) to a pre‑formatted Excel template using smart markers. | Reuse a single Excel file as a template for different datasets by changing the data source name. | Automate production of HR attendance sheets in .NET applications.
// AI Prompts: Add a column that counts "Present" entries for each employee after the smart markers are processed. | Show how to rename the data source from "Attendance" to another name while keeping the same marker syntax. | Provide code to format the Date column as "MM/dd/yyyy" after WorkbookDesigner processing.

using System;
using System.Data;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerMatrixExample
{
    // Demonstrates how to generate an attendance spreadsheet by creating a workbook, adding a header row, inserting smart markers with the [0] index to repeat rows for each DataTable record, naming the marker range "_CellsSmartMarkers", binding the DataTable to a WorkbookDesigner, processing the named range, and saving the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Set up header row for the attendance matrix
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Employee 1");
                cells["C1"].PutValue("Employee 2");
                cells["D1"].PutValue("Employee 3");

                // 3. Insert smart markers using index syntax.
                //    The [0] index tells the designer to repeat the row for each record in the data source.
                cells["A2"].PutValue("&=Attendance.Date[0]");
                cells["B2"].PutValue("&=Attendance.Emp1[0]");
                cells["C2"].PutValue("&=Attendance.Emp2[0]");
                cells["D2"].PutValue("&=Attendance.Emp3[0]");

                // 4. Define the range that contains the smart markers and name it as required by the designer.
                //    The name \"_CellsSmartMarkers\" signals that this range should be processed.
                AsposeRange smartRange = cells.CreateRange("A2:D2");
                smartRange.Name = "_CellsSmartMarkers";

                // 5. Prepare a DataTable that represents the attendance records.
                DataTable attendanceTable = new DataTable("Attendance");
                attendanceTable.Columns.Add("Date", typeof(DateTime));
                attendanceTable.Columns.Add("Emp1", typeof(string));
                attendanceTable.Columns.Add("Emp2", typeof(string));
                attendanceTable.Columns.Add("Emp3", typeof(string));

                // Sample data: three days of attendance
                attendanceTable.Rows.Add(new DateTime(2023, 9, 1), "Present", "Absent", "Present");
                attendanceTable.Rows.Add(new DateTime(2023, 9, 2), "Absent", "Present", "Present");
                attendanceTable.Rows.Add(new DateTime(2023, 9, 3), "Present", "Present", "Absent");

                // 6. Create a WorkbookDesigner, assign the workbook and set the data source.
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Attendance", attendanceTable);

                // 7. Process only the defined range (true = preserve unrecognized markers, not needed here).
                designer.Process(smartRange, true);

                // 8. Save the resulting workbook.
                string outputPath = "AttendanceMatrix.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
