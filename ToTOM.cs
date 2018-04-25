using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using TOMtoAMO.Parsing;
using AMO = Microsoft.AnalysisServices;
using TOM = Microsoft.AnalysisServices.Tabular;
namespace TOMtoAMO
{
    public static class ToTOM
    {


        /* Complete feature list, to be noted when complete support added
         *  - Database
         *      - Direct Query
         *      - Datasources (Done)
         *      - Tables
         *          - Translation of table (Done)
         *          - Attributes
         *              - Translation of Attribute
         *          - Hierarchies
         *              - Translation of Hierarchies
         *              - Levels
         *                  - Translation of Levels
         *      - Measures
         *          - Translation of Measures
         *          - KPI's
         *      - Perspectives
         *      - Roles
         *          - Row Level Security
         *          - Members
         *      - Relationships
         */
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
                TOMTable.IsHidden = !AMODatabase.Cubes[0].Dimensions.FindByName(Dimension.Name).Visible;

                foreach (AMO.Translation AMOTranslation in Dimension.Translations)
                    TranslationHelper.AddTOMTranslation(TOMDatabase, TOMTable, AMOTranslation);
                #region Columns
                foreach (AMO.DimensionAttribute Attribute in Dimension.Attributes)
                {
                    if (Attribute.Type != AMO.AttributeType.RowNumber)
                    {
                        TOM.Column TOMColumn;
                        if (Attribute.NameColumn.Source is AMO.ExpressionBinding)
                        {
                            //Calculated column specific attributes
                            TOM.CalculatedColumn CalculatedColumn = new TOM.CalculatedColumn();
                            CalculatedColumn.Expression = ((AMO.ExpressionBinding)Attribute.NameColumn.Source).Expression;

                            //Set as TOMColumn so generic properties can be applied later
                            TOMColumn = CalculatedColumn;
                        }
                        else
                        {
                            //Data column specific attributes
                            TOM.DataColumn DataColumn = new TOM.DataColumn();
                            DataColumn.SourceColumn = ((AMO.ColumnBinding)Attribute.NameColumn.Source).ColumnID;
                            DataColumn.IsKey = Attribute.Usage == AMO.AttributeUsage.Key;
                            DataColumn.IsNullable = Attribute.KeyColumns[0].NullProcessing != AMO.NullProcessing.Error;

                            //Set as TOMColumn so generic properties can be applied later
                            TOMColumn = DataColumn;
                        }
                        
                        //Generic Properties
                        TOMColumn.Name = Attribute.Name;
                        TOMColumn.Description = Attribute.Description;
                        TOMColumn.DisplayFolder = Attribute.AttributeHierarchyDisplayFolder;
                        TOMColumn.FormatString = Attribute.FormatString;
                        TOMColumn.IsHidden = !Attribute.AttributeHierarchyVisible;
                        TOMColumn.DataType = DataTypeHelper.ToTOMDataType(Attribute.KeyColumns[0].DataType);

                        //Add translations
                        foreach (AMO.Translation AMOTranslation in Attribute.Translations)
                            TranslationHelper.AddTOMTranslation(TOMDatabase, TOMColumn, AMOTranslation);

                        //Finally, add the Column to the Table
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
                #endregion
                TOMDatabase.Model.Tables.Add(TOMTable);
                #region Hierarchies
                foreach (AMO.Hierarchy AMOHierarchy in Dimension.Hierarchies)
                {
                    TOM.Hierarchy TOMHierarchy = new TOM.Hierarchy();
                    TOMHierarchy.Name = AMOHierarchy.Name;
                    TOMHierarchy.Description = AMOHierarchy.Description;
                    TOMHierarchy.DisplayFolder = AMOHierarchy.DisplayFolder;
                    TOMHierarchy.IsHidden = false;

                    //Add translations
                    foreach (AMO.Translation AMOTranslation in AMOHierarchy.Translations)
                        TranslationHelper.AddTOMTranslation(TOMDatabase, TOMHierarchy, AMOTranslation);

                    foreach (AMO.Level AMOLevel in AMOHierarchy.Levels)
                    {
                        TOM.Level TOMLevel = new TOM.Level();
                        TOMLevel.Name = AMOLevel.Name;
                        TOMLevel.Description = AMOLevel.Description;
                        TOMLevel.Column = TOMTable.Columns[AMOLevel.SourceAttribute.Name];
                        TOMHierarchy.Levels.Add(TOMLevel);

                        //Add translations
                        foreach (AMO.Translation AMOTranslation in AMOLevel.Translations)
                            TranslationHelper.AddTOMTranslation(TOMDatabase, TOMLevel, AMOTranslation);
                    }
                    TOMTable.Hierarchies.Add(TOMHierarchy);
                }
                #endregion
                #region Partitions
                foreach(AMO.Partition AMOPartition in AMODatabase.Cubes[0].MeasureGroups.GetByName(TOMTable.Name).Partitions)
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

            #region Measures
            foreach (AMO.Command Command in AMODatabase.Cubes[0].MdxScripts[0].Commands)
            {
                List<MDXCommand> Commands = MDXCommandParser.GetCommands(Command.Text);
                if (Commands.Count > 0)
                {
                    MDXCommand MainCommand = Commands[0];
                    if (MainCommand.Type == CommandType.CreateMeasure)
                    {
                        List<MDXString> Strings = MDXCommandParser.GetStrings(MainCommand.LHS);
                        //Throw exception if we do not have a valid CREATE MEASURE command.
                        if (Strings.Count < 2 || Strings[0].Type != StringType.SingleQuote || Strings[1].Type != StringType.SquareBracket)
                            throw new System.Exception("A CREATE MEASURE statement must at least have two delimited elements (one table, one measure name)");

                        //First, single quoted string, is table name
                        string TableName = Strings[0].Text;
                        //Then, square-bracket delimited string is the measure name.
                        string MeasureName = Strings[1].Text;

                        AMO.CalculationProperty CalculationProperty = AMODatabase.Cubes[0].MdxScripts[0].CalculationProperties.Find("[" + MeasureName.Replace("]", "]]") + "]");
                        TOM.Measure TOMMeasure = new TOM.Measure
                        {
                            Name = MeasureName,
                            Expression = MainCommand.RHS,
                            Description = CalculationProperty?.Description,
                            DisplayFolder = CalculationProperty?.DisplayFolder,
                            FormatString = CalculationProperty?.FormatString.Substring(1, CalculationProperty.FormatString.Length - 2),
                            IsHidden = CalculationProperty == null ? true : !CalculationProperty.Visible
                        };
                        TOMDatabase.Model.Tables[TableName].Measures.Add(TOMMeasure);

                        //Add Translations
                        if (CalculationProperty != null)
                            foreach(AMO.Translation AMOTranslation in CalculationProperty.Translations) 
                                TranslationHelper.AddTOMTranslation(TOMDatabase, TOMMeasure, AMOTranslation);
                                
                    }
                }
            }
            #endregion

            TOMDatabase.Model.Annotations.Add(new TOM.Annotation { Name = "TabularEditor_CompatibilityVersion", Value = AMODatabase.CompatibilityLevel.ToString() });
            //TODO: Handle KPIs
            return TOMDatabase;
        }
    }
}
