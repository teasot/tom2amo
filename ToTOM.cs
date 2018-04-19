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

            #region DataSources
            foreach(AMO.DataSource AMODataSource in AMODatabase.DataSources)
            {
                TOM.ProviderDataSource TOMDataSource = new TOM.ProviderDataSource
                {
                    Description = AMODataSource.Description,
                    ConnectionString = AMODataSource.ConnectionString,
                    MaxConnections = AMODataSource.MaxActiveConnections,
                    Name = AMODataSource.Name,
                    Provider = AMODataSource.ManagedProvider,
                    Timeout = (int)AMODataSource.Timeout.TotalSeconds
                };
                switch (AMODataSource.ImpersonationInfo.ImpersonationMode)
                {
                    case AMO.ImpersonationMode.Default:
                        TOMDataSource.ImpersonationMode = TOM.ImpersonationMode.Default;
                        break;
                    case AMO.ImpersonationMode.ImpersonateAccount:
                        TOMDataSource.ImpersonationMode = TOM.ImpersonationMode.ImpersonateAccount;
                        break;
                    case AMO.ImpersonationMode.ImpersonateAnonymous:
                        TOMDataSource.ImpersonationMode = TOM.ImpersonationMode.ImpersonateAnonymous;
                        break;
                    case AMO.ImpersonationMode.ImpersonateCurrentUser:
                        TOMDataSource.ImpersonationMode = TOM.ImpersonationMode.ImpersonateCurrentUser;
                        break;
                    case AMO.ImpersonationMode.ImpersonateServiceAccount:
                        TOMDataSource.ImpersonationMode = TOM.ImpersonationMode.ImpersonateServiceAccount;
                        break;
                    case AMO.ImpersonationMode.ImpersonateUnattendedAccount:
                        TOMDataSource.ImpersonationMode = TOM.ImpersonationMode.ImpersonateUnattendedAccount;
                        break;
                }

                switch (AMODataSource.Isolation)
                {
                    case AMO.DataSourceIsolation.ReadCommitted:
                        TOMDataSource.Isolation = TOM.DatasourceIsolation.ReadCommitted;
                        break;
                    case AMO.DataSourceIsolation.Snapshot:
                        TOMDataSource.Isolation = TOM.DatasourceIsolation.Snapshot;
                        break;
                }
                TOMDatabase.Model.DataSources.Add(TOMDataSource);
            }
            #endregion

            foreach (AMO.Dimension Dimension in AMODatabase.Dimensions)
            {
                TOM.Table TOMTable = new TOM.Table();
                TOMTable.Name = Dimension.Name;
                #region Columns
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
                            CalculatedColumn.DataType = DataTypeHelper.ToTOMDataType(Attribute.KeyColumns[0].DataType);
                            CalculatedColumn.Description = Attribute.Description;
                            CalculatedColumn.DisplayFolder = Attribute.AttributeHierarchyDisplayFolder;
                            CalculatedColumn.FormatString = Attribute.FormatString;
                            CalculatedColumn.IsHidden = !Attribute.AttributeHierarchyVisible;
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
                //Add sort by columns
                foreach (AMO.DimensionAttribute Attribute in Dimension.Attributes)
                {
                    if (Attribute.Type != AMO.AttributeType.RowNumber && Attribute.OrderByAttribute != null)
                    {
                        TOMTable.Columns[Attribute.Name].SortByColumn = TOMTable.Columns[Attribute.OrderByAttribute.Name];
                    }
                }
                TOMDatabase.Model.Tables.Add(TOMTable);
                #endregion
                #region Hierarchies
                foreach(AMO.Hierarchy AMOHierarchy in Dimension.Hierarchies)
                {
                    TOM.Hierarchy TOMHierarchy = new TOM.Hierarchy();
                    TOMHierarchy.Name = AMOHierarchy.Name;
                    TOMHierarchy.Description = AMOHierarchy.Description;
                    TOMHierarchy.DisplayFolder = AMOHierarchy.DisplayFolder;
                    TOMHierarchy.IsHidden = false;
                    foreach(AMO.Level AMOLevel in AMOHierarchy.Levels)
                    {
                        TOM.Level TOMLevel = new TOM.Level();
                        TOMLevel.Name = AMOLevel.Name;
                        TOMLevel.Description = AMOLevel.Description;
                        TOMLevel.Column = TOMTable.Columns[AMOLevel.SourceAttribute.Name];
                        TOMHierarchy.Levels.Add(TOMLevel);
                    }
                    TOMTable.Hierarchies.Add(TOMHierarchy);
                }
                #endregion
                #region Partitions
                foreach(AMO.Partition AMOPartition in AMODatabase.Cubes[0].MeasureGroups[TOMTable.Name].Partitions)
                {
                    TOM.Partition TOMPartition = new TOM.Partition();
                    TOMPartition.Source = new TOM.QueryPartitionSource
                    {
                        DataSource = TOMDatabase.Model.DataSources[AMOPartition.DataSource.Name],
                        Query = ((AMO.QueryBinding)AMOPartition.Source).QueryDefinition
                    };
                    TOMTable.Partitions.Add(TOMPartition);
                }
                #endregion
            }

            return TOMDatabase;
        }
    }
}
