using AMO = Microsoft.AnalysisServices;
namespace TOM2AMO.File
{
    public static class AMOFile
    {
        /// <summary>
        /// Reads an 1103 or 1100 Database file from the provided file location
        /// </summary>
        /// <param name="FileLocation">The location of the file to be loaded. Usually a .bim file.</param>
        /// <returns></returns>
        public static AMO.Database ReadFromFile(string FileLocation)
        {
            using (System.Xml.XmlReader Reader = System.Xml.XmlReader.Create(FileLocation))
            {
                Reader.ReadToFollowing("Database");
                return (AMO.Database)AMO.Utils.Deserialize(Reader, new AMO.Database());
            }
        }

        /// <summary>
        /// Writes an AMO database to the specified location
        /// </summary>
        /// <param name="FileLocation">The location to write the file to</param>
        /// <param name="Database">The database to serialise</param>
        public static void WriteToFile(string FileLocation, AMO.Database Database)
        {
            System.IO.File.WriteAllText(FileLocation, GetSerialisedXMLString(Database));
        }
        
        public static string GetSerialisedXMLString(AMO.Database AMODatabase)
        {
            using (var StringWriter = new System.IO.StringWriter())
            {
                using (var XMLWriter = System.Xml.XmlWriter.Create(StringWriter, new System.Xml.XmlWriterSettings{
                    Indent = true,
                    IndentChars = "\t",
                    NewLineOnAttributes = true
                }))
                {
                    AMO.Utils.Serialize(XMLWriter, AMODatabase, false);
                }
                return StringWriter.ToString();
            }
        }
    }
}
