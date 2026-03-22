using System;
using System.Text;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

namespace TimelineToHtmlDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and get the first worksheet
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 2. Populate worksheet with sample chronological data
            // ------------------------------------------------------------
            // Header row
            cells["A1"].Value = "Date";
            cells["B1"].Value = "Amount";

            // Sample rows
            DateTime[] dates = new DateTime[]
            {
                new DateTime(2023, 1, 5),
                new DateTime(2023, 2, 12),
                new DateTime(2023, 3, 20),
                new DateTime(2023, 4, 15),
                new DateTime(2023, 5, 30)
            };

            int[] amounts = new int[] { 120, 250, 180, 300, 210 };

            for (int i = 0; i < dates.Length; i++)
            {
                cells[i + 1, 0].Value = dates[i];
                cells[i + 1, 1].Value = amounts[i];
            }

            // Apply a date style to the first column for proper display
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "m/d/yyyy";
            for (int i = 1; i <= dates.Length; i++)
            {
                cells[i, 0].SetStyle(dateStyle);
            }

            // ------------------------------------------------------------
            // 3. Create a PivotTable that will serve as the data source for the Timeline
            // ------------------------------------------------------------
            PivotTableCollection pivots = sheet.PivotTables;
            // The source range includes the header row and all data rows
            int pivotIndex = pivots.Add("A1:B6", "D1", "SalesPivot");
            PivotTable pivot = pivots[pivotIndex];
            // Add the Date field to the Row area (this will be the timeline base field)
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            // Add the Amount field to the Data area
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            // Refresh to calculate the pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // 4. Add a Timeline linked to the PivotTable
            // ------------------------------------------------------------
            // Place the Timeline starting at cell G1 (row 0, column 6)
            int timelineIndex = sheet.Timelines.Add(pivot, 0, 6, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];

            // Optional: customize the Timeline appearance using the Shape object
            timeline.Shape.Width = 600;   // width in pixels
            timeline.Shape.Height = 80;   // height in pixels
            timeline.Caption = "Sales Timeline";

            // ------------------------------------------------------------
            // 5. Generate HTML markup representing the chronological timeline
            // ------------------------------------------------------------
            // We'll build a simple horizontal timeline using an unordered list.
            // Each list item will display the date and the corresponding amount.
            StringBuilder htmlBuilder = new StringBuilder();

            htmlBuilder.AppendLine("<!DOCTYPE html>");
            htmlBuilder.AppendLine("<html lang=\"en\">");
            htmlBuilder.AppendLine("<head>");
            htmlBuilder.AppendLine("    <meta charset=\"UTF-8\">");
            htmlBuilder.AppendLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            htmlBuilder.AppendLine("    <title>Chronological Timeline</title>");
            // Basic CSS for a horizontal timeline
            htmlBuilder.AppendLine("    <style>");
            htmlBuilder.AppendLine("        .timeline {");
            htmlBuilder.AppendLine("            display: flex;");
            htmlBuilder.AppendLine("            overflow-x: auto;");
            htmlBuilder.AppendLine("            padding: 20px 0;");
            htmlBuilder.AppendLine("            list-style: none;");
            htmlBuilder.AppendLine("            border-top: 2px solid #3498db;");
            htmlBuilder.AppendLine("        }");
            htmlBuilder.AppendLine("        .timeline li {");
            htmlBuilder.AppendLine("            position: relative;");
            htmlBuilder.AppendLine("            flex: 0 0 auto;");
            htmlBuilder.AppendLine("            margin-right: 40px;");
            htmlBuilder.AppendLine("            text-align: center;");
            htmlBuilder.AppendLine("        }");
            htmlBuilder.AppendLine("        .timeline li::before {");
            htmlBuilder.AppendLine("            content: '';");
            htmlBuilder.AppendLine("            position: absolute;");
            htmlBuilder.AppendLine("            top: -12px;");
            htmlBuilder.AppendLine("            left: 50%;");
            htmlBuilder.AppendLine("            transform: translateX(-50%);");
            htmlBuilder.AppendLine("            width: 12px;");
            htmlBuilder.AppendLine("            height: 12px;");
            htmlBuilder.AppendLine("            background: #3498db;");
            htmlBuilder.AppendLine("            border-radius: 50%;");
            htmlBuilder.AppendLine("        }");
            htmlBuilder.AppendLine("        .timeline .date {");
            htmlBuilder.AppendLine("            font-weight: bold;");
            htmlBuilder.AppendLine("            margin-bottom: 5px;");
            htmlBuilder.AppendLine("        }");
            htmlBuilder.AppendLine("        .timeline .value {");
            htmlBuilder.AppendLine("            color: #555;");
            htmlBuilder.AppendLine("        }");
            htmlBuilder.AppendLine("    </style>");
            htmlBuilder.AppendLine("</head>");
            htmlBuilder.AppendLine("<body>");
            htmlBuilder.AppendLine($"    <h2>{timeline.Caption}</h2>");
            htmlBuilder.AppendLine("    <ul class=\"timeline\">");

            // Loop through the original data rows to build timeline entries
            for (int i = 0; i < dates.Length; i++)
            {
                string dateStr = dates[i].ToString("MMM dd, yyyy");
                string amountStr = amounts[i].ToString("C0"); // currency format without decimals

                htmlBuilder.AppendLine("        <li>");
                htmlBuilder.AppendLine($"            <div class=\"date\">{dateStr}</div>");
                htmlBuilder.AppendLine($"            <div class=\"value\">{amountStr}</div>");
                htmlBuilder.AppendLine("        </li>");
            }

            htmlBuilder.AppendLine("    </ul>");
            htmlBuilder.AppendLine("</body>");
            htmlBuilder.AppendLine("</html>");

            // ------------------------------------------------------------
            // 6. Write the HTML to a file for dynamic integration
            // ------------------------------------------------------------
            string htmlPath = Path.Combine(Environment.CurrentDirectory, "timeline.html");
            File.WriteAllText(htmlPath, htmlBuilder.ToString(), Encoding.UTF8);

            // ------------------------------------------------------------
            // 7. Save the workbook (optional, demonstrates lifecycle rule usage)
            // ------------------------------------------------------------
            string workbookPath = Path.Combine(Environment.CurrentDirectory, "TimelineDemo.xlsx");
            workbook.Save(workbookPath); // save workbook with the embedded Timeline control

            Console.WriteLine("HTML timeline generated at: " + htmlPath);
            Console.WriteLine("Workbook saved at: " + workbookPath);
        }
    }
}