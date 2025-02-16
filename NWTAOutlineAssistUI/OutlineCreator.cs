using NWTAOutlineAssist;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Diagnostics.CodeAnalysis;
using Microsoft.VisualBasic.FileIO;

namespace NWTAOutlineAssistUI
{
    public class OutlineCreator
    {
        private OAConfiguration _config;
        
        public OutlineCreator(OAConfiguration config)
        {
            _config = config;
        }

        public void CreateOutline()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _config.RoleAssignments = _config.OutlineName.Trim() + " Role Assignments.xlsx";

            var sourceTemplate = AppDomain.CurrentDomain.BaseDirectory + @"\Templates\NWTARoleAssignmentTemplate.xlsx";
            var outlineTemplate = AppDomain.CurrentDomain.BaseDirectory + @"\Templates\Output\" + _config.OutlineTemplate;
            var ssName = _config.FullPath(_config.RoleAssignments);
            var ssRoster = _config.FullPath(_config.StaffRoster);
            var ssRoles = _config.FullPath(_config.RoleRequests);
            var ssTemplate = _config.FullPath(_config.OutlineTemplate);

            if (File.Exists(ssName))
            {
                File.Delete(ssName);
            }
            File.Copy(sourceTemplate, ssName);

            if (File.Exists(ssTemplate))
            {
                File.Delete(ssTemplate);
            }
            File.Copy(outlineTemplate, ssTemplate);

            Dictionary<string, string> roleMappings = ReadRoleMappings();
            List<StaffMan> staffMen;
            if (ssRoster.EndsWith(".csv", StringComparison.InvariantCultureIgnoreCase))
                staffMen = ReadStaffNamesCSV(ssRoster);
            else
                staffMen = ReadStaffNames(ssRoster);

            // sort the list by the number of staffings
            staffMen.Sort((PosA, PosB) => PosB.Staffings.CompareTo(PosA.Staffings));

            if (ssRoles.EndsWith(".csv", StringComparison.InvariantCultureIgnoreCase))
                ReadRoleRequestsCSV(ssRoles, staffMen, roleMappings);
            else
                ReadRoleRequests(ssRoles, staffMen, roleMappings);

            PopulateRoleAssignments(ssName, staffMen);
        }

        Dictionary<string, string> ReadRoleMappings()
        {
            var mappings = new Dictionary<string, string>();
            var lines = File.ReadLines(AppDomain.CurrentDomain.BaseDirectory + @"Templates\RoleMappings.csv");
            foreach (var line in lines)
            {
                string[] values = line.Split(',');
                mappings[values[0]] = values[1];
            }
            return mappings;
        }

        List<StaffMan> ReadStaffNames(string ssRoster)
        {
            try
            {
                List<StaffMan> StaffList = new List<StaffMan>();
                using (var package = new ExcelPackage(new FileInfo(ssRoster)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    for (int row = 2; ; ++row)
                    {
                        if (worksheet.Cells[row, 1].Value != null)
                        {
                            var name = worksheet.Cells[row, 1].Value.ToString().Trim();
                            var ldrTrk = worksheet.Cells[row, 6].Value != null ? worksheet.Cells[row, 6].Value.ToString() : null;
                            var elder = worksheet.Cells[row, 7].Value != null ? worksheet.Cells[row, 7].Value.ToString() : null;

                            var staffMan = new StaffMan();
                            staffMan.Name = name;
                            staffMan.Staffings = int.Parse(worksheet.Cells[row, 5].Value.ToString());
                            staffMan.Role = OutlineData.TranslateRole(ldrTrk, elder);
                            StaffList.Add(staffMan);
                        }
                        else
                            break;
                    }
                }
                return StaffList;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred processing the staff roster spreadsheet", ex);
            }
        }

        List<StaffMan> ReadStaffNamesCSV(string ssRoster)
        {
            try
            {
                List<StaffMan> StaffList = new List<StaffMan>();
                int row = 0;
                using (TextFieldParser parser = new TextFieldParser(ssRoster))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    while (!parser.EndOfData)
                    {
                        string[] values = parser.ReadFields();
                        if (row++ == 0)
                            continue;

                        var name = values[0];
                        var ldrTrk = String.IsNullOrWhiteSpace(values[5]) ? null : values[5];
                        var elder = String.IsNullOrWhiteSpace(values[6]) ? null : values[6];

                        var staffMan = new StaffMan();
                        staffMan.Name = name;
                        staffMan.Staffings = int.Parse(values[4]);
                        staffMan.Role = OutlineData.TranslateRole(ldrTrk, elder);
                        StaffList.Add(staffMan);
                    }
                }
                return StaffList;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred processing the staff roster spreadsheet", ex);
            }
        }


        void ReadRoleRequests(string ssRoles, List<StaffMan> staffMen, Dictionary<string, string> roleMappings)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(ssRoles)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    for (int row = 1; ; ++row)
                    {
                        if (worksheet.Cells[row, 1].Value != null)
                        {
                            var name = worksheet.Cells[row, 1].Value.ToString().Trim();
                            var role = worksheet.Cells[row, 2].Value.ToString().Trim();
                            StaffMan staffMan = staffMen.Find(x => x.Name == name);
                            if (staffMan != null)
                            {
                                string roleId = null;
                                roleMappings.TryGetValue(role, out roleId);
                                if (roleId != null)
                                    staffMan.ReqRoles.Add(roleId);
                                else
                                    System.Console.WriteLine("ReadRoleRequests: {0} not assigned for {1}", role, name);
                            }
                        }
                        else
                            break;
                    }

                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred processing the role requests spreadsheet", ex);
            }
        }

        void ReadRoleRequestsCSV(string ssRoles, List<StaffMan> staffMen, Dictionary<string, string> roleMappings)
        {
            try
            {
                int row = 0;
                using (TextFieldParser parser = new TextFieldParser(ssRoles))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");
                    while (!parser.EndOfData)
                    {
                        string[] values = parser.ReadFields();
                        if (row++ == 0)
                            continue;

                        var name = values[0];
                        var role = values[1];
                        StaffMan staffMan = staffMen.Find(x => x.Name == name);
                        if (staffMan != null)
                        {
                            string roleId = null;
                            roleMappings.TryGetValue(role, out roleId);
                            if (roleId != null)
                                staffMan.ReqRoles.Add(roleId);
                            else
                                System.Console.WriteLine("ReadRoleRequests: {0} not assigned for {1}", role, name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred processing the role requests spreadsheet", ex);
            }
        }


        void PopulateRoleAssignments(string ssName, List<StaffMan> staffMen)
        {
            using (var package = new ExcelPackage(new FileInfo(ssName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                Dictionary<string, int> role2rowMap = new Dictionary<string, int>();
                for (int rowIndex = 1; ;++rowIndex)
                {
                    if (worksheet.Cells[rowIndex, 1].Value != null)
                    {
                        var processId = worksheet.Cells[rowIndex, 1].Value.ToString().Trim();
                        if (processId == "END")
                            break;

                        if (!String.IsNullOrWhiteSpace(processId))
                        {
                            try { role2rowMap[processId] = rowIndex; } catch (Exception ex) { throw new ApplicationException($"The RoleAssignmentTemplate duplicates the RoleId {processId} in row {rowIndex}", ex); }
                        }
                    }
                }

                int colIndex = 3;
                foreach (StaffMan staffMan in staffMen)
                {
                    worksheet.Cells[1, colIndex].Value = staffMan.Name;
                    worksheet.Cells[2, colIndex].Value = staffMan.Staffings;
                    if (staffMan.Role != null)
                    {
                        worksheet.Cells[3, colIndex].Value = staffMan.Role;
                    }
                    foreach (string req in staffMan.ReqRoles)
                    {
                        if (req != "0")
                        {
                            var cell = worksheet.Cells[role2rowMap[req], colIndex];
                            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                        }
                    }
                    colIndex++;
                }
                package.Save();
            }
        }
    }
}
