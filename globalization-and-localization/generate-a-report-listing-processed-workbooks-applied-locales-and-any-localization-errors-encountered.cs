using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookLocalizationReport
{
    // Custom globalization settings that can be extended for different locales
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        private readonly string _locale;

        public CustomGlobalizationSettings(string locale)
        {
            _locale = locale;
        }

        // Example: localize boolean values (can be expanded per locale)
        public override string GetBooleanValueString(bool bv)
        {
            return _locale switch
            {
                "zh" => bv ? "真" : "假",
                "ru" => bv ? "ИСТИНА" : "ЛОЖЬ",
                _ => base.GetBooleanValueString(bv)
            };
        }

        // Example: localize error strings (can be expanded per locale)
        public override string GetErrorValueString(string err)
        {
            return _locale switch
            {
                "zh" => err switch
                {
                    "#NAME?" => "#名称?",
                    "#DIV/0!" => "#除以0!",
                    "#REF!" => "#引用!",
                    "#VALUE!" => "#值!",
                    "#N/A" => "#无效!",
                    "#NUM!" => "#数值!",
                    "#NULL!" => "#空!",
                    _ => err
                },
                "ru" => err switch
                {
                    "#NAME?" => "#ИМЯ?",
                    "#DIV/0!" => "#ДЕЛ/0!",
                    "#REF!" => "#ССЫЛКА!",
                    "#VALUE!" => "#ЗНАЧ?",
                    "#N/A" => "#Н/Д",
                    "#NUM!" => "#ЧИСЛО!",
                    "#NULL!" => "#ПУСТО!",
                    _ => err
                },
                _ => base.GetErrorValueString(err)
            };
        }
    }

    class Program
    {
        static void Main()
        {
            // List of workbooks to process (full paths)
            var workbookPaths = new List<string>
            {
                @"C:\Data\Report_en.xlsx",
                @"C:\Data\Report_zh.xlsx",
                @"C:\Data\Report_ru.xlsx"
            };

            // Mapping of workbook path to locale identifier
            var workbookLocales = new Dictionary<string, string>
            {
                { @"C:\Data\Report_en.xlsx", "en" },
                { @"C:\Data\Report_zh.xlsx", "zh" },
                { @"C:\Data\Report_ru.xlsx", "ru" }
            };

            // Store report entries
            var reportLines = new List<string>();
            reportLines.Add("Processed Workbooks Report");
            reportLines.Add("==========================");
            reportLines.Add("");

            foreach (var path in workbookPaths)
            {
                // Verify file exists
                if (!File.Exists(path))
                {
                    reportLines.Add($"File not found: {path}");
                    continue;
                }

                // Load workbook (lifecycle rule: use constructor with path)
                var workbook = new Workbook(path);

                // Determine locale for this workbook
                workbookLocales.TryGetValue(path, out string locale);
                locale ??= "en";

                // Apply custom globalization settings (lifecycle rule: set property)
                workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings(locale);

                // Collect localization errors
                var errors = new List<string>();

                // Iterate through all worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    // Define the used range to avoid scanning empty cells
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            // Check if the cell contains an error value
                            if (cell.Type == CellValueType.IsError)
                            {
                                // Retrieve the original error string
                                string originalError = cell.StringValue;
                                // Get the localized representation via globalization settings
                                string localizedError = workbook.Settings.GlobalizationSettings.GetErrorValueString(originalError);
                                errors.Add($"Sheet '{sheet.Name}'!{cell.Name}: {originalError} => {localizedError}");
                            }
                        }
                    }
                }

                // Add entry to report
                reportLines.Add($"Workbook: {Path.GetFileName(path)}");
                reportLines.Add($"Locale applied: {locale}");
                if (errors.Count == 0)
                {
                    reportLines.Add("No localization errors detected.");
                }
                else
                {
                    reportLines.Add("Localization errors:");
                    foreach (var err in errors)
                    {
                        reportLines.Add("  - " + err);
                    }
                }
                reportLines.Add("");

                // (Optional) Save the workbook after applying globalization settings
                // Using lifecycle rule: Save method
                string outputPath = Path.Combine(Path.GetDirectoryName(path) ?? "", $"Processed_{Path.GetFileName(path)}");
                workbook.Save(outputPath);
            }

            // Output the report to console
            foreach (var line in reportLines)
            {
                Console.WriteLine(line);
            }

            // Also write the report to a text file
            string reportFile = @"C:\Data\WorkbookLocalizationReport.txt";
            File.WriteAllLines(reportFile, reportLines);
            Console.WriteLine($"Report saved to: {reportFile}");
        }
    }
}