// Title: Generate an Excel attendance matrix with Aspose.Cells smart markers using index syntax in C#
// AI Prompts: Write C# code that creates a workbook, defines a named range, and uses &[Attendance][${row}][col] smart markers to fill employee attendance data from a two‑dimensional array. | Show how to configure WorkbookDesigner to bind an object[,] as the "Attendance" data source and process the smart markers for a matrix layout. | Demonstrate adding a calculated column that counts present days per employee using a smart‑marker expression in the same worksheet.
// Common Searches: how to use Aspose.Cells smart markers with index syntax to populate a table in C# | c# Aspose.Cells create attendance sheet from 2d array | named range _CellsSmartMarkers Aspose.Cells example | populate Excel matrix using WorkbookDesigner and smart markers | Aspose.Cells smart marker row index placeholder ${row} usage
// Tags: Aspose.Cells smart markers index syntax | C# populate Excel matrix from 2D array | WorkbookDesigner data source object array | named range _CellsSmartMarkers for smart markers | attendance table generation with Aspose.Cells

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerMatrixDemo
{
    // The sample creates a new workbook, adds header cells for an attendance matrix, inserts smart markers using the &[Attendance][${row}][col] index syntax, defines the named range "_CellsSmartMarkers" to limit processing, prepares a two‑dimensional object array containing employee names and daily attendance booleans, binds this array to the "Attendance" marker via WorkbookDesigner, processes all smart markers, and saves the filled worksheet as AttendanceMatrix.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();                     // create workbook
                Worksheet sheet = workbook.Worksheets[0];               // get first sheet

                // 2. Set up header row for the attendance matrix
                sheet.Cells["A1"].PutValue("Employee");
                sheet.Cells["B1"].PutValue("Day 1");
                sheet.Cells["C1"].PutValue("Day 2");
                sheet.Cells["D1"].PutValue("Day 3");

                // 3. Insert smart markers that use index syntax.
                //    The syntax &[Attendance][${row}][colIndex] will be replaced by the
                //    corresponding element from the 2‑dimensional data source.
                //    Row index (${row}) is automatically increased for each repeated row.
                sheet.Cells["A2"].PutValue("&=[Attendance][${row}][0]"); // Employee name
                sheet.Cells["B2"].PutValue("&=[Attendance][${row}][1]"); // Day 1 attendance
                sheet.Cells["C2"].PutValue("&=[Attendance][${row}][2]"); // Day 2 attendance
                sheet.Cells["D2"].PutValue("&=[Attendance][${row}][3]"); // Day 3 attendance

                // 4. Define the range that contains the smart markers.
                //    Naming the range "_CellsSmartMarkers" tells the designer to process only this range.
                AsposeRange smartRange = sheet.Cells.CreateRange("A2:D2");
                smartRange.Name = "_CellsSmartMarkers";

                // 5. Prepare a 2‑dimensional array with attendance data.
                //    First column = employee name, subsequent columns = attendance (true/false).
                object[,] attendanceData = new object[,]
                {
                    { "John Doe", true,  false, true  },
                    { "Jane Smith", false, true,  true  },
                    { "Bob Lee",   true,  true,  false }
                };

                // 6. Set up the WorkbookDesigner, assign the data source, and process the smart markers.
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                    // LineByLine is obsolete; using named range smart markers eliminates the need.
                };
                designer.SetDataSource("Attendance", attendanceData);
                designer.Process();    // process all smart markers in the workbook

                // 7. Save the resulting workbook.
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
