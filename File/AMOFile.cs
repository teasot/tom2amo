using AMO = Microsoft.AnalysisServices;
namespace TOMtoAMO.File
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
            using (System.Xml.XmlWriter Writer = System.Xml.XmlWriter.Create(FileLocation))
            {
                AMO.Utils.Serialize(Writer, Database, false);
            }
        }

    }
}
