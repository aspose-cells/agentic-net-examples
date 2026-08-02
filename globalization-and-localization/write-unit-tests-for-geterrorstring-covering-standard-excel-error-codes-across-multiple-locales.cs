using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Custom globalization settings that map standard Excel error strings to friendly messages.
    public class CustomErrorGlobalizationSettings : SettableGlobalizationSettings
    {
        public override string GetErrorValueString(string err)
        {
            return err switch
            {
                "#DIV/0!" => "Division by zero",
                "#N/A" => "Not available",
                "#VALUE!" => "Invalid value",
                "#NAME?" => "Invalid name",
                "#REF!" => "Invalid reference",
                "#NUM!" => "Invalid number",
                "#NULL!" => "Null intersection",
                "#SPILL!" => "Spill error",
                "#CALC!" => "Calculation error",
                "#CONNECT!" => "Connection error",
                "#BUSY!" => "Busy error",
                "#BLOCKED!" => "Blocked error",
                "#UNKNOWN!" => "Unknown error",
                "#TIMEOUT!" => "Timeout error",
                "#EXTERNAL!" => "External error",
                "#FIELD!" => "Field error",
                _ => base.GetErrorValueString(err)
            };
        }
    }

    class Program
    {
        // List of standard Excel error strings.
        private static readonly string[] StandardErrors = new[]
        {
            "#DIV/0!", "#N/A", "#VALUE!", "#NAME?", "#REF!", "#NUM!", "#NULL!",
            "#SPILL!", "#CALC!", "#CONNECT!", "#BUSY!", "#BLOCKED!", "#UNKNOWN!",
            "#TIMEOUT!", "#EXTERNAL!", "#FIELD!"
        };

        static void Main()
        {
            RunDefaultGlobalizationSettingsTest();
            RunCustomGlobalizationSettingsTest();
        }

        // Test default globalization settings – should return the same string.
        static void RunDefaultGlobalizationSettingsTest()
        {
            try
            {
                var workbook = new Workbook();
                var settings = workbook.Settings.GlobalizationSettings;

                foreach (var err in StandardErrors)
                {
                    string result = settings.GetErrorValueString(err);
                    if (result != err)
                        Console.WriteLine($"FAIL: Expected {err}, got {result}");
                    else
                        Console.WriteLine($"PASS: {err}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected exception in Default_GlobalizationSettings_Test: {ex.Message}");
            }
        }

        // Test custom globalization settings – should map errors to friendly messages.
        static void RunCustomGlobalizationSettingsTest()
        {
            try
            {
                var workbook = new Workbook();
                var customSettings = new CustomErrorGlobalizationSettings();
                workbook.Settings.GlobalizationSettings = customSettings;

                var expectedMappings = new Dictionary<string, string>
                {
                    {"#DIV/0!", "Division by zero"},
                    {"#N/A", "Not available"},
                    {"#VALUE!", "Invalid value"},
                    {"#NAME?", "Invalid name"},
                    {"#REF!", "Invalid reference"},
                    {"#NUM!", "Invalid number"},
                    {"#NULL!", "Null intersection"},
                    {"#SPILL!", "Spill error"},
                    {"#CALC!", "Calculation error"},
                    {"#CONNECT!", "Connection error"},
                    {"#BUSY!", "Busy error"},
                    {"#BLOCKED!", "Blocked error"},
                    {"#UNKNOWN!", "Unknown error"},
                    {"#TIMEOUT!", "Timeout error"},
                    {"#EXTERNAL!", "External error"},
                    {"#FIELD!", "Field error"}
                };

                foreach (var kvp in expectedMappings)
                {
                    string result = customSettings.GetErrorValueString(kvp.Key);
                    if (result != kvp.Value)
                        Console.WriteLine($"FAIL: {kvp.Key} expected {kvp.Value}, got {result}");
                    else
                        Console.WriteLine($"PASS: {kvp.Key} mapped correctly");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected exception in Custom_GlobalizationSettings_Test: {ex.Message}");
            }
        }
    }
}