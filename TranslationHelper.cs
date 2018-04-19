using AMO = Microsoft.AnalysisServices;
using TOM = Microsoft.AnalysisServices.Tabular;
namespace TOMtoAMO
{
    public static class TranslationHelper
    {
        public static int GetCultureFromName(string Name)
        {
            return new System.Globalization.CultureInfo(Name).LCID;
        }

        public static AMO.Translation GetTranslation(TOM.Culture TOMCulture, TOM.MetadataObject TOMObject)
        {
            AMO.Translation AMOTranslation = new AMO.Translation
            {
                Language = GetCultureFromName(TOMCulture.Name),
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
    }
}
