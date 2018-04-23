using TOM = Microsoft.AnalysisServices.Tabular;
namespace TOMtoAMO.File
{
    public static class TOMFile
    {
        /// <summary>
        /// Reads a TOM Database file from the provided file location
        /// </summary>
        /// <param name="FileLocation">The location of the file to be loaded. Usually a .bim file.</param>
        /// <returns></returns>
        public static TOM.Database ReadFromFile(string FileLocation)
        {
            return TOM.JsonSerializer.DeserializeDatabase(System.IO.File.ReadAllText(FileLocation));
        }

        /// <summary>
        /// Writes a TOM database to the specified location
        /// </summary>
        /// <param name="FileLocation">The location to write the file to</param>
        /// <param name="Database">The database to serialise</param>
        public static void WriteToFile(string FileLocation, TOM.Database Database)
        {
            string SerialisedDatabase = TOM.JsonSerializer.SerializeDatabase(Database, new TOM.SerializeOptions
            {
                IgnoreInferredObjects = true,
                IgnoreInferredProperties = true,
                IgnoreTimestamps = true,
                IncludeRestrictedInformation = false,
                SplitMultilineStrings = true
            });
            System.IO.File.WriteAllText(FileLocation, SerialisedDatabase);
        }
    }
}
