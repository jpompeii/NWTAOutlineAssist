
namespace NWTAOutlineAssist
{
    public class OAConfiguration
    {
        public string OutlineName { get; set; }
        public string OutlineTemplate { get; set; }
        public string StaffRoster { get; set; }
        public string RoleRequests { get; set; }
        public string RoleAssignments { get; set; }
        public string OutlineOutput { get; set; }
        public string OutlineFolder { get; set; }
        public string FullPath(string file)  
        {
            return OutlineFolder + "\\" + file;
        }
    }
}
