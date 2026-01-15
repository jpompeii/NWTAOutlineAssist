using NWTARules;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace NWTAOutlineAssistUI
{
    public interface INWTAAssignments
    {
       Dictionary<string, Function> Functions { get; }
       void ProcessAssignments();
    }

    public class NWTAAssignments : INWTAAssignments
    {
        public Dictionary<int, StaffMan> Staff { get; set; } = new Dictionary<int, StaffMan>();
        public Dictionary<string, Function> Functions { get; set; } = new Dictionary<string, Function>();

        ExcelWorksheet worksheet;
        int lastMan = 0;
        string assignmentSheet;
        TextWriter console;

        public static INWTAAssignments Create(string outlineDir, string assnFile, TextWriter console)
        {
            if (String.IsNullOrWhiteSpace(assnFile))
                throw new ArgumentNullException(nameof(assnFile), "Assignment file cannot be null or empty.");
            if (console == null)
                throw new ArgumentNullException(nameof(console), "Console cannot be null.");

            if (assnFile.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                return new NWTAAssignmentsFromGoogle(assnFile, console);
            else
                return new NWTAAssignments(outlineDir + "\\" + assnFile, console);
        }

        public NWTAAssignments(string assnFile, TextWriter console)
        {
            assignmentSheet = assnFile;
            this.console = console;
        }

        public void ProcessAssignments()
        {
            using (var package = new ExcelPackage(new FileInfo(assignmentSheet)))
            {
                worksheet = package.Workbook.Worksheets[0];
                ReadStaff();
                ReadFunctions();
            }
        }

        void ReadStaff()
        {
            int manIdx = 3;
            while (worksheet.Cells[1, manIdx].Value != null)
            {
                var cell = worksheet.Cells[1, manIdx];
                var staffMan = new StaffMan();
                var nameParts = cell.Value.ToString().Split(',', 2);
                staffMan.Name = nameParts[0];
                staffMan.Staffings = int.Parse(worksheet.Cells[2, manIdx].Value.ToString());
                staffMan.Role = worksheet.Cells[3, manIdx].Value == null ? String.Empty : worksheet.Cells[3, manIdx].Value.ToString();
                Staff[manIdx] = staffMan;
                ++manIdx;
            }
            lastMan = manIdx - 1;
        }

        void ReadFunctions()
        {
            int row = 1;
            do
            {
                var cell = worksheet.Cells[row, 1];
                if (cell.Value != null)
                {
                    var fnId = cell.Value.ToString();
                    if (fnId == "END")
                        break;

                    var name = worksheet.Cells[row, 2].Value.ToString();
                    var funct = new Function(fnId, name);
                    if (fnId.StartsWith("TX"))
                    {
                        funct.Staff.Add(new NameAndRole(worksheet.Cells[row, 2].Value.ToString(), ""));
                    }
                    else
                    {
                        for (int col = 3; col <= lastMan; col++)
                        {
                            cell = worksheet.Cells[row, col];
                            if (cell.Value != null && !String.IsNullOrWhiteSpace(cell.Value.ToString()))
                            {
                                name = Staff[col].Name;
                                var role = TranslateRole(cell.Value.ToString());
                                if (role != null)
                                    funct.Staff.Add(new NameAndRole(name, role));
                                else
                                    console.WriteLine("ReadFunctions: value assigned: {0} in cell {1},{2} is ignored", cell.Value.ToString(), fnId, col);
                            }
                        }
                    }
                    Functions[fnId] = funct;
                }
                ++row;

            } while (true);

        }

        public static string TranslateRole(string assn)
        {
            // move this to a configuration
            if (assn == "x" || assn == "#" || assn == "y" || assn == "B" || assn == "X" || assn == "b")
                return "team";
            else if (assn == "L" || assn == "l")
                return "leader";
            else if (assn == "C" || assn == "c")
                return "colead";
            else if (assn == "M" || assn == "m")
                return "team";
            else if (assn == "a" || assn == "A")
                return "assist";
            else return null;
        }

    }
}
