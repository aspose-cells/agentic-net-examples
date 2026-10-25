// Title: Generate an attendance matrix in an Excel file using indexed smart markers with Aspose.Cells for C#
// Common Searches: how to fill a multi‑day attendance table in Excel using Aspose.Cells smart markers with index syntax | C# Aspose.Cells WorkbookDesigner populate boolean array columns from a list of objects | using indexed smart markers to map array elements to separate Excel columns in .NET | create attendance report Excel template and bind it to a DataSet with Aspose.Cells | Aspose.Cells smart marker example for dynamic column generation based on array length
// Tags: smart marker indexing Aspose.Cells C# | populate Excel matrix from boolean array | WorkbookDesigner DataSet binding | attendance report generation Excel | dynamic column mapping with smart markers

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AttendanceSmartMarkerDemo
{
    // Represents attendance for a single day.
    // The sample loads an Excel template that contains indexed smart markers, builds a DataSet with an Attendance table (Day and separate Present columns), converts a list of Attendance objects into rows, assigns the DataSet to WorkbookDesigner, processes the smart markers to expand the matrix, and saves the populated report as a new Excel file.
    public class Attendance
    {
        // Day label (e.g., "2023-09-01").
        public string Day { get; set; } = string.Empty;

        // Attendance status for each employee (true = present, false = absent).
        public bool[] Present { get; set; } = Array.Empty<bool>();
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "AttendanceTemplate.xlsx";
                const string outputPath = "AttendanceReport.xlsx";

                // Verify that the template file exists to avoid FileNotFoundException.
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the Excel template that contains smart markers.
                Workbook workbook = new Workbook(templatePath);

                // Prepare sample attendance data.
                var attendanceData = new List<Attendance>
                {
                    new Attendance { Day = "2023-09-01", Present = new bool[] { true, false, true } },
                    new Attendance { Day = "2023-09-02", Present = new bool[] { true, true, true } },
                    new Attendance { Day = "2023-09-03", Present = new bool[] { false, true, false } },
                    new Attendance { Day = "2023-09-04", Present = new bool[] { true, true, false } },
                    new Attendance { Day = "2023-09-05", Present = new bool[] { false, false, true } }
                };

                // Convert the list to a DataSet because WorkbookDesigner.SetDataSource expects a DataSet.
                DataSet dataSet = new DataSet();
                DataTable attendanceTable = new DataTable("Attendance");

                // Define columns: Day and three presence columns (adjust count as needed).
                attendanceTable.Columns.Add("Day", typeof(string));
                attendanceTable.Columns.Add("Present0", typeof(bool));
                attendanceTable.Columns.Add("Present1", typeof(bool));
                attendanceTable.Columns.Add("Present2", typeof(bool));

                // Populate rows.
                foreach (var a in attendanceData)
                {
                    DataRow row = attendanceTable.NewRow();
                    row["Day"] = a.Day;
                    row["Present0"] = a.Present.Length > 0 ? a.Present[0] : false;
                    row["Present1"] = a.Present.Length > 1 ? a.Present[1] : false;
                    row["Present2"] = a.Present.Length > 2 ? a.Present[2] : false;
                    attendanceTable.Rows.Add(row);
                }

                dataSet.Tables.Add(attendanceTable);

                // Process the smart markers using WorkbookDesigner.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSet);
                designer.Process();

                // Save the filled workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
