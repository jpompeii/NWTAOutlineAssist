using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NWTAOutlineAssistUI
{
    public class OutlineData
    {
        public static Dictionary<string, ProcessRole> RoleMap = new Dictionary<string, ProcessRole>()
        {
            { "x",  ProcessRole.Staff },
            { "y",  ProcessRole.NoRole },
            { "B",  ProcessRole.Staff },
            { "L",  ProcessRole.Leader },
            { "M",  ProcessRole.Medic },
            { "m",  ProcessRole.Music },
            { "C",  ProcessRole.CoLead },
            { "CC",  ProcessRole.CLC },
        };

        public static Dictionary<ProcessRole, string> RoleStrings = new Dictionary<ProcessRole, string>()
        {
            { ProcessRole.Staff, "" },
            { ProcessRole.Leader, "Leader" },
            { ProcessRole.Medic, "Medic" },
            { ProcessRole.Music, "Music" },
            { ProcessRole.CoLead, "Co-Lead" },
            { ProcessRole.CLC, "CLC" },
        };

        public static string TranslateRole(string ldrTrk, string elder)
        {
            if (elder != null && elder == "Ritual Elder")
                return "RE";
            if (ldrTrk == null)
                return null;
            else if (ldrTrk == "L")
                return "FL";
            else if (ldrTrk == "CL")
                return "CL";
            else if (ldrTrk == "CLC")
                return "CC";
            else if (ldrTrk == "LIT")
                return "L";
            else
                return null;
        }
    }
}
