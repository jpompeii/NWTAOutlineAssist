using NWTARules;
using System.Collections.Generic;
using System.IO;
using GoogleSheetsWrapper;
using System;
using System.IO.Compression;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NWTAOutlineAssistUI
{
    class NWTAAssignmentsFromGoogle : INWTAAssignments
    {
        SheetHelper sheet;
        TextWriter console;
        string assnUrl;
        int lastMan = 0;
        Dictionary<int, StaffMan> Staff = new Dictionary<int, StaffMan>();

        public NWTAAssignmentsFromGoogle(string assnUrl, TextWriter console)
        {
            this.console = console;
            this.assnUrl = assnUrl;
        }

        public Dictionary<string, Function> Functions { get; set; } = new Dictionary<string, Function>();

        void Init()
        {
            var serviceAccount = "service@nwta-outline-assist.iam.gserviceaccount.com";
            var jsonZip = AppDomain.CurrentDomain.BaseDirectory + @"Etc\Config.zip";
            string jsonCredsContent;
            string documentId = null;

            using (FileStream zipToOpen = new FileStream(jsonZip, FileMode.Open, FileAccess.Read))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Read))
                {
                    ZipArchiveEntry keyEntry = archive.GetEntry("key.txt");
                    using (StreamReader reader = new StreamReader(keyEntry.Open()))
                    {
                        jsonCredsContent = reader.ReadToEnd();
                    }
                }
            }

            Uri u;
            if (!Uri.TryCreate(assnUrl, UriKind.Absolute, out u))
            {
                throw new Exception("Invalid URL for assignment sheet in google drive");
            }

            if (!u.Host.StartsWith(@"docs.google.com"))
            {
                throw new Exception("Invalid URL for assignment sheet in google drive");
            }
            string[] pathElems = u.AbsolutePath.Split('/');
            foreach (var elem in pathElems)
            {
                if (elem == "d")
                {
                    // next element is the document ID
                    int idx = Array.IndexOf(pathElems, elem);
                    if (idx + 1 < pathElems.Length)
                    {
                        documentId = pathElems[idx + 1];
                    }
                }
            }
            if (string.IsNullOrEmpty(documentId))
            {
                throw new Exception("Cannot find document ID in google drive URL for assignment sheet");
            }

            sheet = new SheetHelper(documentId, serviceAccount, "Staff Roles");
            sheet.Init(jsonCredsContent);
        }

        public void ProcessAssignments()
        {
            Init();
            var sheetData = sheet.GetRows(new SheetRange("", 1, 1, 60));
            ReadStaff(sheetData);
            ReadFunctions(sheetData);
        }

        void ReadStaff(IList<IList<object>> sheetData)
        {
            IList<object> nameRow = sheetData[0];
            IList<object> staffingsRow = sheetData[1];
            IList<object> roleRow = sheetData[2];

            for (int i = 0; i < nameRow.Count; ++i)
            {
                if (i < 2)
                    continue;

                if (nameRow[i] == null || String.IsNullOrWhiteSpace(nameRow[i].ToString()))
                {
                    lastMan = i - 1;
                    break;
                }

                var staffMan = new StaffMan();
                var nameParts = nameRow[i].ToString().Split(',', 2);
                staffMan.Name = nameParts[0];
                staffMan.Staffings = int.Parse(staffingsRow[i].ToString());
                if (roleRow.Count <= i)
                    staffMan.Role = String.Empty;
                else
                    staffMan.Role = roleRow[i] == null ? String.Empty : roleRow[i].ToString();
                
                Staff[i] = staffMan;
            }
        }

        void ReadFunctions(IList<IList<object>> sheetData)
        {
            for (int rowIdx = 5; rowIdx < sheetData.Count;)
            {
                var row = sheetData[rowIdx];
                if (row.Count == 0)
                {
                    rowIdx++;
                    continue;   
                }

                var cell = row[0];
                if (cell != null)
                {
                    var fnId = cell.ToString();
                    if (fnId == "END")
                        break;

                    var name = row[1].ToString();
                    var funct = new Function(fnId, name);
                    if (fnId.StartsWith("TX"))
                    {
                        funct.Staff.Add(new NameAndRole(name, ""));
                    }
                    else
                    {
                        for (int col = 2; col <= lastMan && col < row.Count; col++)
                        {
                            cell = row[col];
                            if (cell != null && !String.IsNullOrWhiteSpace(cell.ToString()))
                            {
                                name = Staff[col].Name;
                                var role = NWTAAssignments.TranslateRole(cell.ToString());
                                if (role != null)
                                    funct.Staff.Add(new NameAndRole(name, role));
                                else
                                    console.WriteLine("ReadFunctions: value assigned: {0} in cell {1},{2} is ignored", cell.ToString(), fnId, col);
                            }
                        }
                    }
                    Functions[fnId] = funct;
                }
                ++rowIdx;

            } 
        }

    }
}