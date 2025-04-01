using NWTARules;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace NWTAOutlineAssistUI
{
    public class NWTAAssignments
    {
        public Dictionary<int, StaffMan> Staff { get; set; } = new Dictionary<int, StaffMan>();
        public Dictionary<string, Function> Functions { get; set; } = new Dictionary<string, Function>();

        ExcelWorksheet worksheet;
        int lastMan = 0;
        string assignmentSheet;
        TextWriter console;

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

        public void ReadStaff()
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

        public void ReadFunctions()
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

        public string TranslateRole(string assn)
        {
            // move this to a configuration
            if (assn == "x" | assn == "y" || assn == "B")
                return "team";
            else if (assn == "L")
                return "leader";
            else if (assn == "C")
                return "colead";
            else if (assn == "M")
                return "team";
            else if (assn == "a")
                return "assist";
            else return null;
        }

    }
}
