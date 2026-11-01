// Title: Generate a project timeline Excel file from a custom template and replace {ProjectName} placeholders using Aspose.Cells for .NET (C#)
// Common Searches: how to create a timeline chart in Excel using Aspose.Cells C# | replace placeholder {ProjectName} in an Excel workbook with Aspose.Cells .NET | load an existing Excel template and add a stacked bar chart with Aspose.Cells C# | populate project start and end dates into an Excel worksheet programmatically | save a generated timeline report to a specific folder using Aspose.Cells C#
// Tags: Aspose.Cells load template workbook | Aspose.Cells add stacked bar chart | Aspose.Cells replace cell placeholder text | Aspose.Cells write project data to Excel | Aspose.Cells save workbook to file

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads a custom Excel template (or creates a new workbook), writes project headers and rows, adds a stacked bar chart that visualizes start and end dates, scans all cells to replace the {ProjectName} placeholder with actual project names, and saves the resulting timeline report to the specified output file.
class TimelineGenerator
{
    static void Main()
    {
        try
        {
            // Paths for template and output files
            string templatePath = "Template.xlsx";
            string outputPath = "TimelineReport.xlsx";

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            Workbook workbook;

            // Load existing template if it exists; otherwise create a new workbook
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Timeline";
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Sample project data
            var projects = new List<Project>
            {
                new Project { Name = "Project Alpha", Start = new DateTime(2023, 1, 1), End = new DateTime(2023, 3, 15) },
                new Project { Name = "Project Beta",  Start = new DateTime(2023, 2, 10), End = new DateTime(2023, 5, 30) },
                new Project { Name = "Project Gamma", Start = new DateTime(2023, 4, 1), End = new DateTime(2023, 8, 20) }
            };

            // Insert headers starting at A2 (row index 1)
            int headerRow = 1;
            sheet.Cells[headerRow, 0].PutValue("Project");
            sheet.Cells[headerRow, 1].PutValue("Start Date");
            sheet.Cells[headerRow, 2].PutValue("End Date");

            // Populate project rows
            for (int i = 0; i < projects.Count; i++)
            {
                int row = headerRow + i + 1;
                sheet.Cells[row, 0].PutValue(projects[i].Name);
                sheet.Cells[row, 1].PutValue(projects[i].Start);
                sheet.Cells[row, 2].PutValue(projects[i].End);
            }

            // Add a stacked bar chart (rows 6‑26, columns A‑J)
            try
            {
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 25, 9);
                Chart chart = sheet.Charts[chartIndex];

                // Define series range (Start and End dates)
                string seriesRange = $"B{headerRow + 2}:C{headerRow + projects.Count + 1}";
                chart.NSeries.Add(seriesRange, true);

                // Define category (project names)
                chart.NSeries.CategoryData = $"A{headerRow + 2}:A{headerRow + projects.Count + 1}";
                chart.Title.Text = "Project Timeline";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Chart creation error: {ex.Message}");
            }

            // Replace placeholder text {ProjectName} with actual name (example uses first project)
            try
            {
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;
                for (int r = 0; r <= maxRow; r++)
                {
                    for (int c = 0; c <= maxCol; c++)
                    {
                        var cell = sheet.Cells[r, c];
                        if (cell.Type == CellValueType.IsString && cell.StringValue.Contains("{ProjectName}"))
                        {
                            cell.PutValue(cell.StringValue.Replace("{ProjectName}", projects[0].Name));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Placeholder replacement error: {ex.Message}");
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple class to hold project information
    class Project
    {
        public string Name { get; set; } = string.Empty;
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
    }
}
