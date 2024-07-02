using NWTAOutlineAssist;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;

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
            File.Copy(outlineTemplate, ssTemplate);

            Dictionary<string, int> roleMappings = ReadRoleMappings();
            List<StaffMan> staffMen = ReadStaffNames(ssRoster);
            staffMen.Sort((PosA, PosB) => PosB.Staffings.CompareTo(PosA.Staffings));

            foreach (StaffMan man in staffMen)
            {
                System.Console.Out.WriteLine("name: " + man.Name + " staffings: " + man.Staffings);
            }

            ReadRoleRequests(ssRoles, staffMen, roleMappings);
            PopulateRoleAssignments(ssName, staffMen);
        }

        Dictionary<string, int> ReadRoleMappings()
        {
            var mappings = new Dictionary<string, int>();
            var lines = File.ReadLines(AppDomain.CurrentDomain.BaseDirectory + @"Templates\RoleMappings.csv");
            foreach (var line in lines)
            {
                string[] values = line.Split(',');
                int process = int.Parse(values[1]);
                mappings[values[0]] = process;
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
                            var name = worksheet.Cells[row, 1].Value.ToString();
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

        void ReadRoleRequests(string ssRoles, List<StaffMan> staffMen, Dictionary<string, int> roleMappings)
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
                            var name = worksheet.Cells[row, 1].Value.ToString();
                            var role = worksheet.Cells[row, 2].Value.ToString();
                            StaffMan staffMan = staffMen.Find(x => x.Name == name);
                            if (staffMan != null)
                            {
                                int roleId = 0;
                                roleMappings.TryGetValue(role, out roleId);
                                if (roleId > 0)
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

        void PopulateRoleAssignments(string ssName, List<StaffMan> staffMen)
        {
            using (var package = new ExcelPackage(new FileInfo(ssName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colIndex = 2;
                foreach (StaffMan staffMan in staffMen)
                {
                    worksheet.Cells[1, colIndex].Value = staffMan.Name;
                    worksheet.Cells[2, colIndex].Value = staffMan.Staffings;
                    if (staffMan.Role != null)
                    {
                        worksheet.Cells[3, colIndex].Value = staffMan.Role;
                    }
                    foreach (int req in staffMan.ReqRoles)
                    {
                        var cell = worksheet.Cells[req, colIndex];
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                    }
                    colIndex++;
                }
                package.Save();
            }
        }
    }
}
