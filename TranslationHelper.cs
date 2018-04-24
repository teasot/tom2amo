using AMO = Microsoft.AnalysisServices;
using TOM = Microsoft.AnalysisServices.Tabular;
namespace TOMtoAMO
{
    public static class TranslationHelper
    {
        public static int GetLCIDFromCultureName(string Name)
        {
            return new System.Globalization.CultureInfo(Name).LCID;
        }

        public static string GetCultureNameFromLCID(int LCID)
        {
            return new System.Globalization.CultureInfo(LCID).Name;

        }

        public static AMO.Translation GetTranslation(TOM.Culture TOMCulture, TOM.MetadataObject TOMObject)
        {
            AMO.Translation AMOTranslation = new AMO.Translation
            {
                Language = GetLCIDFromCultureName(TOMCulture.Name),
                Caption = TOMCulture.ObjectTranslations[TOMObject, TOM.TranslatedProperty.Caption]?.Value,
                Description = TOMCulture.ObjectTranslations[TOMObject, TOM.TranslatedProperty.Description]?.Value,
                DisplayFolder = TOMCulture.ObjectTranslations[TOMObject, TOM.TranslatedProperty.DisplayFolder]?.Value
            };

            // If translation has no properties, it does not exist. 
            // As such, return null unless it has at least one valid property
            if (AMOTranslation.Caption != null || AMOTranslation.Description != null || AMOTranslation.DisplayFolder != null)
                return AMOTranslation;
            else
                return null;
        }

        /// <summary>
        /// Adds a translation for the TOM object based on the supplied AMOTranslation. 
        /// Creates the relevant culture if it does not exist.
        /// </summary>
        /// <param name="TOMDatabase">The TOM Database to add the Translation to</param>
        /// <param name="TOMObject">The TOM Object the translation is being added for</param>
        /// <param name="AMOTranslation">The AMO Translation being transferred to the TOM Database</param>
        public static void AddTOMTranslation(TOM.Database TOMDatabase, TOM.MetadataObject TOMObject, AMO.Translation AMOTranslation)
        {
            string CultureName = TranslationHelper.GetCultureNameFromLCID(AMOTranslation.Language);

            //Add Culture if it does not exist
            if (!TOMDatabase.Model.Cultures.ContainsName(CultureName))
                TOMDatabase.Model.Cultures.Add(new TOM.Culture { Name = CultureName });

            //Get existing culture
            TOM.Culture TOMCulture = TOMDatabase.Model.Cultures.Find(CultureName);

            //Add the various translated properties to the translation
            TOMCulture.ObjectTranslations.Add(
                new TOM.ObjectTranslation { Object = TOMObject, Property = TOM.TranslatedProperty.Caption, Value = AMOTranslation.Caption }
            );
            TOMCulture.ObjectTranslations.Add(
                new TOM.ObjectTranslation { Object = TOMObject, Property = TOM.TranslatedProperty.Description, Value = AMOTranslation.Description }
            );
            TOMCulture.ObjectTranslations.Add(
                new TOM.ObjectTranslation { Object = TOMObject, Property = TOM.TranslatedProperty.DisplayFolder, Value = AMOTranslation.DisplayFolder }
            );
        }
    }
}
