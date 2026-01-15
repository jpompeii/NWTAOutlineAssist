using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualBasic.FileIO;
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
            if (rosterFile.EndsWith(".csv", StringComparison.InvariantCultureIgnoreCase))
                ReadStaffCSV();
            else
                ReadStaffXls();

            return StaffList;
        }

        void ReadStaffXls()
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
        }

        void ReadStaffCSV()
        {
            try
            {
                int row = 0;
                using (TextFieldParser parser = new TextFieldParser(rosterFile))
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
                        staffMan.Name = values[0];
                        staffMan.Area = OutlinePrint.CellString(values[1]);
                        staffMan.Community = OutlinePrint.CellString(values[2]);
                        staffMan.WarriorName = OutlinePrint.CellString(values[3]);
                        staffMan.Staffings = int.Parse(OutlinePrint.CellString(values[4]));
                        staffMan.Role = OutlinePrint.CellString(values[5]);
                        staffMan.Elder = OutlinePrint.CellString(values[6]);
                        staffMan.Email = OutlinePrint.CellString(values[7]);
                        staffMan.Phone = OutlinePrint.CellString(values[8]);
                        staffMan.City = OutlinePrint.CellString(values[10]);
                        staffMan.State = OutlinePrint.CellString(values[11]);
                        staffMan.CPR = OutlinePrint.CellString(values[13]);
                        StaffList.Add(staffMan);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred processing the staff roster CSV", ex);
            }
        }

    }
}
