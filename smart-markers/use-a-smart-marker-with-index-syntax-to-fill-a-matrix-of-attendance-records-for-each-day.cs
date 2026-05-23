using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerAttendanceDemo
{
    // Data model representing an employee and his/her attendance for a week
    public class Employee
    {
        // Initialize to avoid non‑nullable warnings
        public string Name { get; set; } = string.Empty;
        public bool[] Days { get; set; } = Array.Empty<bool>();
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // ---------- Create a new workbook (create rule) ----------
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                var cells = ws.Cells;

                // ---------- Set up the template with smart markers ----------
                // Header row
                cells["A1"].PutValue("Employee");
                cells["B1"].PutValue("Day1");
                cells["C1"].PutValue("Day2");
                cells["D1"].PutValue("Day3");
                cells["E1"].PutValue("Day4");
                cells["F1"].PutValue("Day5");

                // Data rows start at A2. Use smart markers with index syntax to fill the matrix.
                cells["A2"].PutValue("&=Attendance.Name");
                cells["B2"].PutValue("&=Attendance.Days[0]");
                cells["C2"].PutValue("&=Attendance.Days[1]");
                cells["D2"].PutValue("&=Attendance.Days[2]");
                cells["E2"].PutValue("&=Attendance.Days[3]");
                cells["F2"].PutValue("&=Attendance.Days[4]");

                // Define the range that contains the smart markers and name it "_CellsSmartMarkers"
                // This enables range‑based processing (required when LineByLine = false)
                AsposeRange smartRange = cells.CreateRange("A2:F2");
                smartRange.Name = "_CellsSmartMarkers";

                // ---------- Prepare sample attendance data ----------
                var employees = new List<Employee>
                {
                    new Employee { Name = "Alice", Days = new[] { true,  false, true,  true,  false } },
                    new Employee { Name = "Bob",   Days = new[] { false, true,  true,  false, true  } },
                    new Employee { Name = "Carol", Days = new[] { true,  true,  true,  true,  true  } }
                };

                // ---------- Set up WorkbookDesigner and assign the data source ----------
                var designer = new WorkbookDesigner
                {
                    Workbook = wb
                };
                designer.SetDataSource("Attendance", employees);

                // ---------- Process the smart markers (process rule) ----------
                designer.Process();

                // ---------- Save the populated workbook (save rule) ----------
                const string outputPath = "AttendanceMatrix.xlsx";

                // Ensure the directory exists (prevents FileNotFoundException on save)
                var outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}