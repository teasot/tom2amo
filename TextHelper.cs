using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TOMtoAMO
{
    public static class TextHelper
    {
        private enum CommentType
        {
            block,
            line,
            none
        }
        public static string ParseDAX(string DAX)
        {

            //Split into lines and remove and leftover carriage returns
            List<string> Lines = new List<string>(DAX.Split('\n'));
            var l = DAX.Split('\n');
            for (int i=0; i < Lines.Count; i++)
                Lines[i] = Lines[i].Replace("\r", "");


            int commentlineStart;
            int commentCharStart;
            int commentlineEnd;
            int commentCharEnd;
            string ProcessedDAX = "";
            for (int i = 0; i < Lines.Count; i++)
            {
                if (Lines[i].Length >= 2) {
                    //j represents the scond pointer.
                    //Start one character from the start, end at the end
                    for (int CurIndex = 1; CurIndex < Lines[i].Length; CurIndex++)
                    {
                        int PrevIndex = CurIndex - 1;
                        string currentPair = new string(new char[] { Lines[i][PrevIndex], Lines[i][CurIndex] });
                    }
                }
            }

            string FinalDAX = "";
            return FinalDAX;
        }
    }
}
