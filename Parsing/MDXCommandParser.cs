using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TOM2AMO.Parsing
{
    public static class MDXCommandParser
    {
        private static readonly string[] KEY_WORDS = new string[]{
            "CREATE MEASURE",
            "CREATE KPI CURRENTCUBE",
            "ALTER CUBE CURRENTCUBE UPDATE DIMENSION",
            "CREATE MEMBER CURRENTCUBE.Measures."
        };

        private enum CommentType
        {
            Block,
            Line,
            None
        }


        //Gets a list of commands.
        private static Tuple<List<MDXCommand>, List<MDXString>> GenericParse(string DAX)
        {
            List<MDXCommand> Commands = new List<MDXCommand>();
            List<MDXString> Strings = new List<MDXString>();

            //Split into lines
            List<string> Lines;
            if (DAX.Split('\n').Length == 0)
                Lines = new List<string> { DAX };
            else
                Lines = new List<string>(DAX.Split('\n'));

            //The following is a hack: append ' ' to the end of each line as an "empty" character.
            //This way, later, when I parse it, and stop one from the end, I do not miss anything
            for (int i = 0; i < Lines.Count; i++)
                if (Lines[i].Length > 0)
                    Lines[i] += ' ';

            
            CommentType CurrentCommentType = CommentType.None;
            MDXString CurrentString;
            bool skipNext = false;

            MDXCommand CurrentCommand = null;
            List<string> ProcessedLines = new List<string>();
            for (int i = 0; i < Lines.Count; i++)
            {
                if (CurrentCommand != null && CurrentCommand.RHS != null)
                    CurrentCommand.RHS += System.Environment.NewLine;
                List<char> ProcessedDAX = new List<char>();
                skipNext = false;
                CurrentString = null;
                //Cancel comment if line comment on new line
                if (CurrentCommentType == CommentType.Line)
                    CurrentCommentType = CommentType.None;

                if (Lines[i].Length >= 2)
                {
                    //Start start, end one from the end
                    for (int CurIndex = 0, NextIndex = 1; CurIndex < Lines[i].Length - 1; CurIndex++, NextIndex++)
                    {
                        string currentPair = new string(new char[] { Lines[i][CurIndex], Lines[i][NextIndex] });
                        //Skip the next character if necessary
                        if (CurrentCommentType == CommentType.Line || skipNext)
                        {
                            if(skipNext)
                                skipNext = false;
                            ProcessedDAX.Add(Lines[i][CurIndex]);
                            if (CurrentCommand != null)
                            {
                                CurrentCommand.Text += Lines[i][CurIndex];
                                if (CurrentCommand.RHS == null)
                                    CurrentCommand.LHS += Lines[i][CurIndex];
                                else if (!(Lines[i][CurIndex] == ';' && CurrentCommentType == CommentType.None && CurrentString == null))
                                    CurrentCommand.RHS += Lines[i][CurIndex];
                            }
                            continue;
                        }
                        //Handle string state if potentially entering or exiting string
                        if (Lines[i][CurIndex] == '"' || Lines[i][CurIndex] == '\'' || Lines[i][CurIndex] == '[' || Lines[i][CurIndex] == ']')
                        {
                            //If we are not in the string, enter it, and push the opening quote
                            //Unlike most "strings", square bracket strings have different opening and closing characters
                            if (CurrentString == null && Lines[i][CurIndex] != ']')
                            {
                                switch (Lines[i][CurIndex])
                                {
                                    case '"':
                                        CurrentString = new MDXString(StringType.DoubleQuote);
                                        break;
                                    case '\'':
                                        CurrentString = new MDXString(StringType.SingleQuote);
                                        break;
                                    case '[':
                                        CurrentString = new MDXString(StringType.SquareBracket);
                                        break;
                                    default:
                                        throw new Exception(string.Format("Unhandled string opening character: {0}", Lines[i][CurIndex]));
                                }
                            }
                            else
                            {
                                //Skip if terminating character (skip second double quote), otherwise close of string
                                //Note double single quotes are not escaped for single quote strings in MDX in this particular context
                                if (
                                    (Lines[i][NextIndex] == Lines[i][CurIndex] && Lines[i][CurIndex] == '"' && CurrentString.Type == StringType.DoubleQuote)
                                    || (  Lines[i][NextIndex] == Lines[i][CurIndex] && Lines[i][CurIndex] == ']' && CurrentString.Type == StringType.SquareBracket)
                                )
                                {
                                    CurrentString.Text += Lines[i][CurIndex];
                                    skipNext = true;
                                }
                                else if(CurrentString.Type == StringType.DoubleQuote && Lines[i][CurIndex] == '"'
                                    || CurrentString.Type == StringType.SingleQuote && Lines[i][CurIndex] == '\''
                                        || CurrentString.Type == StringType.SquareBracket && Lines[i][CurIndex] == ']')
                                {
                                    //Add the string to the list, and set the current reference to null
                                    Strings.Add(CurrentString);
                                    CurrentString = null;
                                }
                                else
                                    CurrentString.Text += Lines[i][CurIndex];
                            }
                            if (CurrentCommand != null)
                            {
                                CurrentCommand.Text += Lines[i][CurIndex];

                                if (CurrentCommand.RHS == null)
                                    CurrentCommand.LHS += Lines[i][CurIndex];
                                else if (!(Lines[i][CurIndex] == ';' && CurrentCommentType == CommentType.None && CurrentString == null))
                                    CurrentCommand.RHS += Lines[i][CurIndex];
                            }
                            ProcessedDAX.Add(Lines[i][CurIndex]);
                            continue;
                        }
                        if (CurrentString == null)
                        {
                            //If not currently in a block comment, skip the line
                            if ((currentPair == "--" || currentPair == "//") && CurrentCommentType != CommentType.Block)
                            {
                                CurrentCommentType = CommentType.Line;
                                ProcessedDAX.Add(Lines[i][CurIndex]);
                            }
                            else if (currentPair == "/*" && CurrentCommentType == CommentType.None)
                            {
                                CurrentCommentType = CommentType.Block;
                                skipNext = true;
                            }
                            else if (currentPair == "*/" && CurrentCommentType == CommentType.Block)
                            {
                                CurrentCommentType = CommentType.None;
                                skipNext = true;
                            }
                        }
                        else
                            CurrentString.Text += Lines[i][CurIndex];

                        ProcessedDAX.Add(Lines[i][CurIndex]);
                        if (CurrentCommand != null)
                        {
                            CurrentCommand.Text += Lines[i][CurIndex];
                            if (CurrentCommand.RHS == null)
                                CurrentCommand.LHS += Lines[i][CurIndex];
                            else if (!(Lines[i][CurIndex] == ';' && CurrentCommentType == CommentType.None && CurrentString == null))
                                CurrentCommand.RHS += Lines[i][CurIndex];
                        }
                        if (CurrentCommentType == CommentType.None && CurrentString == null)
                        {
                            string KeyWord = AtKeyWord(ProcessedDAX);
                            if (KeyWord != null)
                                switch (KeyWord)
                                {
                                    case "CREATE MEASURE":
                                        CurrentCommand = new MDXCommand(CommandType.CreateMeasure);
                                        break;
                                    case "CREATE KPI CURRENTCUBE":
                                        CurrentCommand = new MDXCommand(CommandType.CreateKPI);
                                        break;
                                }
                            if(Lines[i][CurIndex] == ';')
                            {

                                if (CurrentCommand != null)
                                {
                                    Commands.Add(CurrentCommand);
                                    CurrentCommand = null;
                                }
                            }
                            if(Lines[i][CurIndex] == '=' && CurrentCommand != null)
                            {
                                CurrentCommand.RHS = "";
                            }
                        }
                    }
                }
                else
                    ProcessedDAX.AddRange(Lines[i]);
                ProcessedLines.Add(new string(ProcessedDAX.ToArray()));
            }
            return new Tuple<List<MDXCommand>, List<MDXString>>(Commands, Strings);
        }

        public static List<MDXCommand> GetCommands(string DAX)
        {
            return GenericParse(DAX).Item1;
        }

        public static List<MDXString> GetStrings(string DAX)
        {
            return GenericParse(DAX).Item2;
        }

        /// <summary>
        /// Given a string of characters, determines if the last set of characters are a key word
        /// </summary>
        /// <param name="CurrentParsedDAXArray"></param>
        /// <returns></returns>
        private static string AtKeyWord(List<char> CurrentParsedDAXArray)
        {
            string CurrentParsedDAX = new string(CurrentParsedDAXArray.ToArray());
            foreach(string KeyWord in KEY_WORDS)
                if(CurrentParsedDAX.Length >= KeyWord.Length)
                    if (CurrentParsedDAX.EndsWith(KeyWord))
                        return KeyWord;
            return null;
        }

    }
}
