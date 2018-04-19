using AMO = Microsoft.AnalysisServices;
using TOM = Microsoft.AnalysisServices.Tabular;
namespace TOMtoAMO
{
    public static class RelationshipHelper
    {
        internal static bool CreateRelationship(TOM.SingleColumnRelationship TOMRelationship, AMO.Database Database)
        {
            AMO.Dimension ToDimension = Database.Dimensions[TOMRelationship.ToTable.Name];
            AMO.DimensionAttribute ToAttribute = ToDimension.Attributes[TOMRelationship.ToColumn.Name];
            AMO.Dimension FromDimension = Database.Dimensions[TOMRelationship.FromTable.Name];
            if (ToDimension.Attributes[TOMRelationship.ToColumn.Name].Usage != AMO.AttributeUsage.Key)
                SetPKColumn(Database, ToAttribute);

            string RelationshipID = System.Guid.NewGuid().ToString();
            
            AMO.Relationship AMORelationship = Database.Dimensions[TOMRelationship.FromTable.Name].Relationships.Add(RelationshipID);

            AMORelationship.FromRelationshipEnd.DimensionID = TOMRelationship.FromTable.Name;
            AMORelationship.FromRelationshipEnd.Attributes.Add(TOMRelationship.FromColumn.Name);
            AMORelationship.FromRelationshipEnd.Multiplicity = AMO.Multiplicity.Many;
            AMORelationship.FromRelationshipEnd.Role = string.Empty;
            AMORelationship.ToRelationshipEnd.DimensionID = TOMRelationship.ToTable.Name;
            AMORelationship.ToRelationshipEnd.Attributes.Add(TOMRelationship.ToColumn.Name);
            AMORelationship.ToRelationshipEnd.Multiplicity = AMO.Multiplicity.One;
            AMORelationship.ToRelationshipEnd.Role = string.Empty;

            if (TOMRelationship.IsActive)
                setActiveRelationship(Database.Cubes[0], TOMRelationship.FromTable.Name, TOMRelationship.FromColumn.Name, TOMRelationship.ToTable.Name, RelationshipID);

            return true;
        }

        private static void SetPKColumn(AMO.Database Database, AMO.DimensionAttribute PKAttribute)
        {
            //Get RowNumber Attribute
            AMO.DimensionAttribute RowNumber = null;
            //Find all 'unwanted' Key attributes, remove their Key definitions and include the attributes in the ["RowNumber"].AttributeRelationships
            foreach (AMO.DimensionAttribute CurrentDimAttribute in PKAttribute.Parent.Attributes)
            {
                foreach (AMO.DimensionAttribute Attribute in PKAttribute.Parent.Attributes)
                    if (Attribute.Type == AMO.AttributeType.RowNumber)
                        RowNumber = Attribute;
                if (RowNumber == null)
                    throw new System.Exception("Unable to find rownumber in Dimension " + PKAttribute.Parent.Name);

                if ((CurrentDimAttribute.Usage == AMO.AttributeUsage.Key) && (CurrentDimAttribute.ID != PKAttribute.ID))
                {
                    CurrentDimAttribute.Usage = AMO.AttributeUsage.Regular;
                    if (CurrentDimAttribute.Type != AMO.AttributeType.RowNumber)
                    {
                        CurrentDimAttribute.KeyColumns[0].NullProcessing = AMO.NullProcessing.Preserve;
                        CurrentDimAttribute.AttributeRelationships.Clear();
                        if (!RowNumber.AttributeRelationships.ContainsName(CurrentDimAttribute.ID))
                        {
                            AMO.AttributeRelationship currentAttributeRelationship = CurrentDimAttribute.Parent.Attributes["RowNumber"].AttributeRelationships.Add(CurrentDimAttribute.ID);
                            currentAttributeRelationship.OverrideBehavior = AMO.OverrideBehavior.None;
                        }
                        RowNumber.AttributeRelationships[CurrentDimAttribute.ID].Cardinality = AMO.Cardinality.Many;
                    }
                }
            }

            //Remove PKColumnName from ["RowNumber"].AttributeRelationships
            int PKAtribRelationshipPosition = RowNumber.AttributeRelationships.IndexOf(PKAttribute.Name);
            if (PKAtribRelationshipPosition != -1)
                RowNumber.AttributeRelationships.RemoveAt(PKAtribRelationshipPosition, true);

            //Define PKColumnName as Key and add ["RowNumber"] to PKColumnName.AttributeRelationships with cardinality of One
            PKAttribute.Usage = AMO.AttributeUsage.Key;
            PKAttribute.KeyColumns[0].NullProcessing = AMO.NullProcessing.Error;
            if (!PKAttribute.AttributeRelationships.ContainsName("RowNumber"))
            {
                AMO.DimensionAttribute currentAttribute = RowNumber;
                AMO.AttributeRelationship currentAttributeRelationship = PKAttribute.AttributeRelationships.Add(currentAttribute.ID);
                currentAttributeRelationship.OverrideBehavior = AMO.OverrideBehavior.None;
            }
            PKAttribute.AttributeRelationships["RowNumber"].Cardinality = AMO.Cardinality.One;
        }
        private static void setActiveRelationship(AMO.Cube currentCube, string MVTableName, string MVColumnName, string PKTableName, string relationshipID)
        {
            AMO.MeasureGroup currentMG = currentCube.MeasureGroups[MVTableName];

            if (!currentMG.Dimensions.Contains(PKTableName))
            {
                AMO.ReferenceMeasureGroupDimension NewReferenceMGDim = new AMO.ReferenceMeasureGroupDimension();
                NewReferenceMGDim.CubeDimensionID = PKTableName;
                NewReferenceMGDim.IntermediateCubeDimensionID = MVTableName;
                NewReferenceMGDim.Materialization = AMO.ReferenceDimensionMaterialization.Regular;
                foreach (AMO.CubeAttribute PKAttribute in currentCube.Dimensions[PKTableName].Attributes)
                {
                    AMO.MeasureGroupAttribute PKMGAttribute = NewReferenceMGDim.Attributes.Add(PKAttribute.AttributeID);
                    System.Data.OleDb.OleDbType PKMGAttributeType = PKAttribute.Attribute.KeyColumns[0].DataType;
                    PKMGAttribute.KeyColumns.Add(new AMO.DataItem(PKTableName, PKAttribute.AttributeID, PKMGAttributeType));
                    PKMGAttribute.KeyColumns[0].Source = new AMO.ColumnBinding(PKTableName, PKAttribute.AttributeID);
                }
                currentMG.Dimensions.Add(NewReferenceMGDim);
            }
            AMO.ReferenceMeasureGroupDimension currentReferenceMGDim = (AMO.ReferenceMeasureGroupDimension)currentMG.Dimensions[PKTableName];
            currentReferenceMGDim.RelationshipID = relationshipID;
            currentReferenceMGDim.IntermediateGranularityAttributeID = MVColumnName;

        }
    }
}
