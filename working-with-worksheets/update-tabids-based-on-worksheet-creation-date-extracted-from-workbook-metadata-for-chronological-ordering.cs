// Title: Update Excel worksheet TabIds chronologically using creation‑date metadata with Aspose.Cells for .NET
// Common Searches: how to assign worksheet TabId based on custom CreationDate property using Aspose.Cells C# | reorder Excel sheet tabs by creation date programmatically with Aspose.Cells | read worksheet custom properties to get creation timestamp in Aspose.Cells .NET | set sequential TabId values after sorting worksheets by metadata in C# | Aspose.Cells example for chronological tab ordering of worksheets
// Tags: assign worksheet TabId by creation date Aspose.Cells | order Excel worksheets using custom property metadata | read custom property CreationDate Aspose.Cells .NET | programmatic chronological tab ordering Aspose.Cells | reassign TabId after sorting worksheets by date

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Properties; // For CustomProperty

namespace UpdateWorksheetTabIdsByCreationDate
{
    // The example loads an Excel workbook, extracts each worksheet's creation date from a custom property named "CreationDate" (falling back to the workbook's built‑in CreatedTime or DateTime.MinValue), sorts the worksheets from oldest to newest, reassigns their TabId values sequentially starting at 1, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Incorrect number of arguments.");
                Console.WriteLine("Usage: UpdateWorksheetTabIdsByCreationDate.exe <inputWorkbookPath> <outputWorkbookPath>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook: {inputPath}");
                Console.WriteLine($"Worksheet count: {workbook.Worksheets.Count}");

                // Collect worksheets together with their creation dates
                var sheetInfoList = new List<(Worksheet sheet, DateTime creationDate)>();

                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    DateTime creationDate = GetWorksheetCreationDate(sheet);
                    sheetInfoList.Add((sheet, creationDate));
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" - Parsed creation date: {creationDate:u}");
                }

                // Order by creation date (oldest first)
                var orderedSheets = sheetInfoList.OrderBy(info => info.creationDate).ToList();

                // Reassign TabId sequentially based on the ordered list
                int tabId = 1;
                foreach (var info in orderedSheets)
                {
                    info.sheet.TabId = tabId;
                    Console.WriteLine($"Assigned TabId {tabId} to worksheet \"{info.sheet.Name}\"");
                    tabId++;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved with updated TabIds to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }

        private static DateTime GetWorksheetCreationDate(Worksheet sheet)
        {
            // Attempt to read a custom property named "CreationDate"
            try
            {
                CustomPropertyCollection customProps = sheet.CustomProperties;
                CustomProperty prop = null;

                // The indexer throws if the property does not exist, so we guard with a try/catch
                try
                {
                    prop = customProps["CreationDate"];
                }
                catch
                {
                    // Property not present
                }

                if (prop != null && !string.IsNullOrEmpty(prop.Value))
                {
                    if (DateTime.TryParse(prop.Value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTime parsed))
                    {
                        return parsed;
                    }
                }
            }
            catch
            {
                // Swallow any exception – fallback will be used
            }

            // Fallback: use the workbook's built‑in creation time (if available)
            try
            {
                Workbook wb = sheet.Workbook;
                if (wb != null && wb.BuiltInDocumentProperties != null)
                {
                    // CreatedTime is a DateTime property
                    DateTime created = wb.BuiltInDocumentProperties.CreatedTime;
                    if (created != DateTime.MinValue)
                    {
                        return created;
                    }
                }
            }
            catch
            {
                // Ignore and fall back further
            }

            // Final fallback: treat as the earliest possible date
            return DateTime.MinValue;
        }
    }
}
