using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TOMtoAMO.Parsing
{
    public class MDXString
    {
        public readonly StringType Type;
        public string Text;
        public MDXString(StringType Type)
        {
            this.Type = Type;
            this.Text = "";
        }
    }
}
