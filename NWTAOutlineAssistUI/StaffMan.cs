using System.Collections.Generic;

namespace NWTAOutlineAssistUI
{
    public class StaffMan
    {
        public string Name;
        public int Staffings;
        public string Role;

        // roster properties
        public string Area;
        public string Community;
        public string WarriorName;
        public string Elder;
        public string Email;
        public string Phone;
        public string City;
        public string State;
        public string CPR;

        public List<int> ReqRoles = new List<int>();
    }

    public enum ProcessRole
    {
        Leader,
        Staff,
        Music,
        Safety,
        Elder,
        Medic,
        CLC,
        CoLead,
        NoRole
    }

    public class AssignmentInfo
    {
        public AssignmentInfo(int procId, ProcessRole role)
        {
            ProcessId = procId;
            Role = role;
        }
        public int ProcessId;
        public ProcessRole Role;
    }

}
