using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AMO = Microsoft.AnalysisServices;
using TOM = Microsoft.AnalysisServices.Tabular;
namespace TOMtoAMO
{
    public static class ToTOM
    {
        public static TOM.Database Database(AMO.Database AMODatabase)
        {
            TOM.Database TOMDatabase = new TOM.Database(AMODatabase.Name);
            TOMDatabase.Model = new TOM.Model();
            TOMDatabase.Model.Name = AMODatabase.Cubes[0].Name;
            foreach(AMO.Dimension Dimension in AMODatabase.Dimensions)
            {
                TOM.Table TOMTable = new TOM.Table();
                TOMTable.Name = Dimension.Name;
                foreach(AMO.DimensionAttribute Attribute in Dimension.Attributes)
                {
                    if (Attribute.Type != AMO.AttributeType.RowNumber)
                    {
                        TOM.Column TOMColumn;
                        if (Attribute.NameColumn.Source is AMO.ExpressionBinding)
                        {
                            TOM.CalculatedColumn CalculatedColumn = new TOM.CalculatedColumn();
                            CalculatedColumn.Name = Attribute.Name;
                            CalculatedColumn.Expression = ((AMO.ExpressionBinding)Attribute.NameColumn.Source).Expression;
                            TOMColumn = CalculatedColumn;
                        }
                        else
                        {
                            TOM.DataColumn DataColumn = new TOM.DataColumn();
                            DataColumn.Name = Attribute.Name;
                            DataColumn.SourceColumn = ((AMO.ColumnBinding)Attribute.NameColumn.Source).ColumnID;
                            DataColumn.DataType = DataTypeHelper.ToTOMDataType(Attribute.KeyColumns[0].DataType);
                            DataColumn.Description = Attribute.Description;
                            DataColumn.DisplayFolder = Attribute.AttributeHierarchyDisplayFolder;
                            DataColumn.FormatString = Attribute.FormatString;
                            DataColumn.IsHidden = !Attribute.AttributeHierarchyVisible;
                            DataColumn.IsKey = Attribute.Usage == AMO.AttributeUsage.Key;
                            DataColumn.IsNullable = Attribute.KeyColumns[0].NullProcessing != AMO.NullProcessing.Error;

                            TOMColumn = DataColumn;
                        }
                        TOMTable.Columns.Add(TOMColumn);
                    }
                }
                TOMDatabase.Model.Tables.Add(TOMTable);
            }

            return TOMDatabase;
        }
    }
}
