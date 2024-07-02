using Microsoft.Extensions.Configuration;
using NWTAOutlineAssist;
using NWTARules;
using OfficeOpenXml;
using System;
using System.IO;
using System.Text;

namespace NWTAOutlineAssistUI
{
    public class OutlinePrint
    {
        public string TemplateFile { get; set; }
        NWTAAssignments assignments;
        StaffRoster roster;
        OAConfiguration configuration;
        TextWriter console;

        public OutlinePrint(OAConfiguration config, TextWriter console)
        {
            this.configuration = config;
            this.console = console;
        }

        public void GenerateOutline()
        {
            configuration.OutlineOutput = configuration.OutlineName.Trim() + " Outline.xlsx";
            TemplateFile = configuration.FullPath(configuration.OutlineOutput);
            if (File.Exists(TemplateFile))
            {
                try
                {
                    File.Delete(TemplateFile);
                }
                catch (Exception ex)
                {
                    throw new ApplicationException($"Cannot overwrite the OutputTemplate: {configuration.OutlineOutput}, do you have it open?", ex);
                }
            }
            File.Copy(configuration.FullPath(configuration.OutlineTemplate), TemplateFile);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            roster = new StaffRoster(configuration.FullPath(configuration.StaffRoster));
            roster.ReadStaff();

            assignments = new NWTAAssignments(configuration.FullPath(configuration.RoleAssignments), console);
            assignments.ProcessAssignments();

            ProcessTemplate();
        }

        public void ProcessTemplate()
        {
            using (var package = new ExcelPackage(new FileInfo(TemplateFile)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[2];
                ProcessCoordinators(worksheet);
                worksheet = package.Workbook.Worksheets[3];
                ProcessOutlineSheet(worksheet);
                worksheet = package.Workbook.Worksheets[0];
                ProcessRoster(worksheet);
                worksheet = package.Workbook.Worksheets[1];
                ProcessStaffPivot(worksheet);
                package.Save();
            }
        }

        public void ProcessRoster(ExcelWorksheet worksheet)
        {
            int row = 2;
            foreach (StaffMan staffMan in roster.StaffList)
            {
                worksheet.Cells[row, 1].Value = staffMan.Name;
                worksheet.Cells[row, 2].Value = staffMan.Area;
                worksheet.Cells[row, 3].Value = staffMan.Community;
                worksheet.Cells[row, 4].Value = staffMan.WarriorName;
                worksheet.Cells[row, 5].Value = staffMan.Staffings;
                worksheet.Cells[row, 6].Value = staffMan.Role;
                worksheet.Cells[row, 7].Value = staffMan.Elder;
                worksheet.Cells[row, 8].Value = staffMan.Email;
                worksheet.Cells[row, 9].Value = staffMan.Phone;
                worksheet.Cells[row, 10].Value = staffMan.City;
                worksheet.Cells[row, 11].Value = staffMan.State;
                worksheet.Cells[row, 12].Value = staffMan.CPR;
                row++;
            }
        }

        public void ProcessStaffPivot(ExcelWorksheet worksheet)
        {
            int[] levelCounts = new int[5];
            int[] levelRows = new int[5];
            int[] levelCols = new int[5];
            for (int i = 0; i < 5; ++i)
                levelRows[i] = 2;

            levelCols[0] = 1;
            levelCols[1] = 4;
            levelCols[2] = 8;
            levelCols[3] = 12;
            levelCols[4] = 16;

            foreach (StaffMan staffMan in roster.StaffList)
            {
                int idx = 4;
                if (staffMan.Role == "L" || staffMan.Role == "CL")
                    idx = 0;
                else if (staffMan.Staffings >= 10)
                    idx = 1;
                else if (staffMan.Staffings >= 5)
                    idx = 2;
                else if (staffMan.Staffings >= 2)
                    idx = 3;

                worksheet.Cells[levelRows[idx], levelCols[idx]].Value = staffMan.Name;
                if (idx == 0)
                    worksheet.Cells[levelRows[idx], levelCols[idx] + 1].Value = staffMan.Role;
                else
                    worksheet.Cells[levelRows[idx], levelCols[idx] + 2].Value = staffMan.Staffings;

                levelCounts[idx]++;
                levelRows[idx]++;
            }
            for (int i = 0; i < 5; i++)
            {
                worksheet.Cells[8 + i, 2].Value = levelRows[i] - 2;
            }
        }

        void ProcessCoordinators(ExcelWorksheet worksheet)
        {
            for (int row = 1; row < 60; ++row)
            {
                for (int col = 1; col <= 5; ++col)
                {
                    if (worksheet.Cells[row, col].Value != null)
                    {
                        var cell = worksheet.Cells[row, col];
                        var textCltn = cell.RichText;
                        foreach (var rtext in textCltn)
                        {
                            var tok = new OToken(rtext.Text);
                            var outString = new StringBuilder();
                            string tokString;
                            bool hasRule = false;
                            while ((tokString = tok.lex()) != null)
                            {
                                if (tok.IsRule)
                                {
                                    hasRule = true;
                                    Rule r = new Rule(tokString);
                                    Function fn = null;
                                    if (assignments.Functions.TryGetValue(r.FunctionID, out fn))
                                        tokString = r.ExpandRule(fn, false);
                                    else
                                        console.WriteLine("Placeholder {0} on row {1} not found", r.FunctionID, row);
                                }
                                outString.Append(tokString);
                            }
                            if (hasRule)
                            {
                                rtext.Text = outString.ToString();
                            }
                        }
                    }

                }
            }
        }

        void ProcessOutlineSheet(ExcelWorksheet worksheet)
        {
            for (int row = 1; row < 250; ++row)
            {
                for (int col = 1; col <= 3; ++col)
                {
                    if (worksheet.Cells[row, col].Value != null)
                    {
                        var cell = worksheet.Cells[row, col];
                        string value = cell.Value.ToString();
                        var textCltn = cell.RichText;
                        foreach (var rtext in textCltn)
                        {
                            var tok = new OToken(rtext.Text);
                            var outString = new StringBuilder();
                            string tokString;
                            bool hasRule = false;
                            while ((tokString = tok.lex()) != null)
                            {
                                if (tok.IsRule)
                                {
                                    hasRule = true;
                                    Rule r = new Rule(tokString);
                                    Function fn = null;
                                    if (assignments.Functions.TryGetValue(r.FunctionID, out fn))
                                        tokString = r.ExpandRule(fn, false);
                                    else
                                        console.WriteLine("Placeholder {0} on row {1} not found", r.FunctionID, row);
                                }
                                outString.Append(tokString);
                            }
                            if (hasRule)
                            {
                                rtext.Text = outString.ToString();
                            }
                        }
                    }

                }
            }
        }

        public static string CellString(ExcelRange cells)
        {
            return cells.Value != null ? cells.Value.ToString() : String.Empty;
        }

    }
}
