using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NWTARules
{
    public class Rule
    {
        public string FunctionID { get; set; }
        public string Role { get; set; } = string.Empty;
        public int ComponentIndex { get; set; } = 0;
        List<RuleOption> Options { get; set; } = new List<RuleOption>();

        // todo: load this from a file
        static public Dictionary<string, string> RoleLabels = new Dictionary<string, string>()
        {
            {"colead", "(CL)"},
            {"music", "(Mu)"},
            {"clc", "(CLC)" },
            {"leader", "(L)"}
        };

        public Rule(string rule) 
        {
            parseRuleText(rule);
        }

        void parseRuleText(string rule)
        {
            string[] parts = rule.Split(':', 2);
            if (parts[0].IndexOf('.') > 0)
            {
                char[] seps = {'[', ']' };
                string[] idParts = parts[0].Split('.', 2);
                FunctionID = idParts[0];
                string[] roleParts = idParts[1].Split(seps, StringSplitOptions.RemoveEmptyEntries);
                Role = roleParts[0];
                if (roleParts.Length == 2)
                {
                    ComponentIndex = int.Parse(roleParts[1]);
                }
                else if (roleParts.Length != 1)
                {
                    throw new ApplicationException("Rule text format error: invalid index format");
                }
            }
            else
                FunctionID = parts[0];

            if (parts.Length > 1)
            {
                // parse the options
                string[] options = parts[1].Split(",");
                foreach (string option in options)
                {
                    string[] optParts = option.Split('=', StringSplitOptions.TrimEntries);
                    if (optParts.Length == 1)
                        Options.Add(new RuleOption(optParts[0], String.Empty));
                    else
                        Options.Add(new RuleOption(optParts[0], optParts[1]));
                }
            }
        }

        public string ExpandRule(Function funct, bool addComma)
        {
            int nameIndex = 1;
            int outCount = 0;
            var outString = new StringBuilder();    
            foreach (NameAndRole nmr in funct.Staff)
            {
                if (String.IsNullOrEmpty(Role) || nmr.Role == "all" || nmr.Role == Role)
                {
                    if (ComponentIndex > 0 && ComponentIndex != nameIndex)
                    {
                        nameIndex++;
                        continue;
                    }
                    if (!String.IsNullOrEmpty(nmr.Name))
                    {
                        if ((TestBooleanRule("comma", true) && outCount == 0) || outCount > 0)
                            outString.Append(", ");
                        
                        outString.Append(nmr.Name);
                        if (TestBooleanRule("label", true))
                            outString.Append(GetRoleLabel(nmr.Role));

                        outCount++;
                        nameIndex++;
                    }
                }
            }
            var result = outString.ToString();
            if (String.IsNullOrEmpty(result) && TestBooleanRule("tbd", true))
                result = "TBD";  // todo: add read/bold highlighting

            return result;
        }

        public bool TestBooleanRule(string name, bool testValue)
        {
            bool val = false;
            RuleOption option = Options.Find(x => x.Name == name);
            if (option != null)
            {
                val = option.IsTrue();
                if (!testValue)
                    val = !val;
            }
            return val;
        }

        public string GetRoleLabel(string name)
        {
            string label = String.Empty;
            RoleLabels.TryGetValue(name, out label);
            return label;
        }
    }

    public class RuleOption
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public RuleOption(string name, string value) 
        {
            Name = name;
            Value = value;
        }

        public bool IsTrue()
        {
            return (String.IsNullOrEmpty(Value) || Value == "true");
        }
    }
}
