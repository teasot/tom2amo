using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TOM2AMO.Parsing
{
    public class MDXCommand
    {
        public CommandType Type;
        public string Text;
        public string LHS;
        public string RHS;

        public MDXCommand(CommandType Type)
        {
            this.Type = Type;
            this.Text = "";
        }
    }
}
