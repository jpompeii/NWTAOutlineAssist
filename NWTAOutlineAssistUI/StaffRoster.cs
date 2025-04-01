using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;


namespace NWTAOutlineAssistUI
{
    public class StaffRoster
    {
        public List<StaffMan> StaffList { get; set; } = new List<StaffMan>();
        string rosterFile;

        public StaffRoster(string rosterFile)
        {
            this.rosterFile = rosterFile;
        }

        public List<StaffMan> ReadStaff()
        {

            using (var package = new ExcelPackage(new FileInfo(rosterFile)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                for (int row = 2; ; ++row)
                {
                    if (worksheet.Cells[row, 1].Value != null)
                    {
                        var staffMan = new StaffMan();
                        staffMan.Name = OutlinePrint.CellString(worksheet.Cells[row, 1]);
                        staffMan.Area = OutlinePrint.CellString(worksheet.Cells[row, 2]);
                        staffMan.Community = OutlinePrint.CellString(worksheet.Cells[row, 3]);
                        staffMan.WarriorName = OutlinePrint.CellString(worksheet.Cells[row, 4]);
                        staffMan.Staffings = int.Parse(OutlinePrint.CellString(worksheet.Cells[row, 5]));
                        staffMan.Role = OutlinePrint.CellString(worksheet.Cells[row, 6]);
                        staffMan.Elder = OutlinePrint.CellString(worksheet.Cells[row, 7]);
                        staffMan.Email = OutlinePrint.CellString(worksheet.Cells[row, 8]);
                        staffMan.Phone = OutlinePrint.CellString(worksheet.Cells[row, 9]);
                        staffMan.City = OutlinePrint.CellString(worksheet.Cells[row, 11]);
                        staffMan.State = OutlinePrint.CellString(worksheet.Cells[row, 12]);
                        staffMan.CPR = OutlinePrint.CellString(worksheet.Cells[row, 14]);
                        StaffList.Add(staffMan);
                    }
                    else
                        break;
                }
            }
            return StaffList;
        }
    }
}
