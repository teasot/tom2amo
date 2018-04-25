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
         *      - Tables (Done)
         *          - Translation of table (Done)
         *          - Columns (Done)
         *              - Translation of Column (Done)
         *          - Hierarchies (Done)
         *              - Translation of Hierarchies (Done)
         *              - Levels (Done)
         *                  - Translation of Levels (Done)
         *      - Measures (Done)
         *          - Translation of Measures (Done)
         *          - KPI's
         *      - Perspectives
         *      - Roles
         *          - Row Level Security
         *          - Members
         *      - Relationships
         */
         /// <summary>
         /// Generates a 1200 Tabular Database, based on the provided 1103 Database.
         /// </summary>
         /// <param name="AMODatabase">The 1103 Database to create 1200 model from</param>
         /// <param name="AddTabularEditorAnnotation">Whether to add an annotation to allow compatibility with TabularEditor</param>
         /// <returns></returns>
        public static TOM.Database Database(AMO.Database AMODatabase, bool AddTabularEditorAnnotation = true)
        {
            //Create the database
            TOM.Database TOMDatabase = new TOM.Database(AMODatabase.Name);

            //Create the model
            TOMDatabase.Model = new TOM.Model();
            TOMDatabase.Model.Name = AMODatabase.Cubes[0].Name;

            #region DataSources
            foreach(AMO.DataSource AMODataSource in AMODatabase.DataSources)
            {
                //Create the data source. We use the ProviderDataSource specifically
                TOM.ProviderDataSource TOMDataSource = new TOM.ProviderDataSource
                {
                    Description = AMODataSource.Description,
                    ConnectionString = AMODataSource.ConnectionString,
                    MaxConnections = AMODataSource.MaxActiveConnections,
                    Name = AMODataSource.Name,
                    Provider = AMODataSource.ManagedProvider,
                    Timeout = (int)AMODataSource.Timeout.TotalSeconds
                };

                //Convert AMO ImpersonationMode enum to TOM ImpersonationMode enum
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

                //Convert AMO Isolation enum to TOM Isolation enum
                switch (AMODataSource.Isolation)
                {
                    case AMO.DataSourceIsolation.ReadCommitted:
                        TOMDataSource.Isolation = TOM.DatasourceIsolation.ReadCommitted;
                        break;
                    case AMO.DataSourceIsolation.Snapshot:
                        TOMDataSource.Isolation = TOM.DatasourceIsolation.Snapshot;
                        break;
                }

                //Add the DataSource
                TOMDatabase.Model.DataSources.Add(TOMDataSource);
            }
            #endregion

            foreach (AMO.Dimension Dimension in AMODatabase.Dimensions)
            {
                //Create the table
                TOM.Table TOMTable = new TOM.Table();
                TOMTable.Description = Dimension.Description;
                TOMTable.Name = Dimension.Name;
                TOMTable.IsHidden = !AMODatabase.Cubes[0].Dimensions.FindByName(Dimension.Name).Visible;

                //Add Translations
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
                            //Create the Calculated Column, and set specific properties
                            TOM.CalculatedColumn CalculatedColumn = new TOM.CalculatedColumn();
                            CalculatedColumn.Expression = ((AMO.ExpressionBinding)Attribute.NameColumn.Source).Expression;

                            //Set as TOMColumn so generic properties can be applied later
                            TOMColumn = CalculatedColumn;
                        }
                        else
                        {
                            //Create the Data Column, and set specific properties
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

                //Add the Table
                TOMDatabase.Model.Tables.Add(TOMTable);
                
                #region Hierarchies
                foreach (AMO.Hierarchy AMOHierarchy in Dimension.Hierarchies)
                {
                    //Create the hierarchy
                    TOM.Hierarchy TOMHierarchy = new TOM.Hierarchy();
                    TOMHierarchy.Name = AMOHierarchy.Name;
                    TOMHierarchy.Description = AMOHierarchy.Description;
                    TOMHierarchy.DisplayFolder = AMOHierarchy.DisplayFolder;

                    //AMO Hierarchies are always visible, from what I can tell
                    TOMHierarchy.IsHidden = false;

                    //Add translations
                    foreach (AMO.Translation AMOTranslation in AMOHierarchy.Translations)
                        TranslationHelper.AddTOMTranslation(TOMDatabase, TOMHierarchy, AMOTranslation);

                    foreach (AMO.Level AMOLevel in AMOHierarchy.Levels)
                    {
                        //Create the level
                        TOM.Level TOMLevel = new TOM.Level();
                        TOMLevel.Name = AMOLevel.Name;
                        TOMLevel.Description = AMOLevel.Description;
                        TOMLevel.Column = TOMTable.Columns[AMOLevel.SourceAttribute.Name];

                        //Add translations
                        foreach (AMO.Translation AMOTranslation in AMOLevel.Translations)
                            TranslationHelper.AddTOMTranslation(TOMDatabase, TOMLevel, AMOTranslation);

                        //Add the Level
                        TOMHierarchy.Levels.Add(TOMLevel);
                    }
                    //Add the Hierarchy
                    TOMTable.Hierarchies.Add(TOMHierarchy);
                }
                #endregion
                #region Partitions
                foreach(AMO.Partition AMOPartition in AMODatabase.Cubes[0].MeasureGroups.GetByName(TOMTable.Name).Partitions)
                {
                    //Create the partition
                    TOM.Partition TOMPartition = new TOM.Partition();
                    TOMPartition.Description = AMOPartition.Description;

                    //Add the query
                    TOMPartition.Source = new TOM.QueryPartitionSource
                    {
                        DataSource = TOMDatabase.Model.DataSources[AMOPartition.DataSource.Name],
                        Query = ((AMO.QueryBinding)AMOPartition.Source).QueryDefinition
                    };

                    //Add the Partition
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

                        //Find the existing CalculationProperty, containing the relevant properties of the Measure, if it exists
                        AMO.CalculationProperty CalculationProperty = AMODatabase.Cubes[0].MdxScripts[0].CalculationProperties.Find("[" + MeasureName.Replace("]", "]]") + "]");

                        //Create the Measure
                        TOM.Measure TOMMeasure = new TOM.Measure
                        {
                            Name = MeasureName,
                            Expression = MainCommand.RHS,
                            Description = CalculationProperty?.Description,
                            DisplayFolder = CalculationProperty?.DisplayFolder,
                            //AMO format string is wrapped in single quotes, so we need to get rid of them here
                            FormatString = CalculationProperty?.FormatString.Substring(1, CalculationProperty.FormatString.Length - 2),
                            IsHidden = CalculationProperty == null ? true : !CalculationProperty.Visible
                        };

                        //Add Translations
                        if (CalculationProperty != null)
                            foreach(AMO.Translation AMOTranslation in CalculationProperty.Translations) 
                                TranslationHelper.AddTOMTranslation(TOMDatabase, TOMMeasure, AMOTranslation);

                        //Add the Measure
                        TOMDatabase.Model.Tables[TableName].Measures.Add(TOMMeasure);

                    }
                }
            }
            #endregion

            #region Perspectives
            foreach(AMO.Perspective AMOPerspective in AMODatabase.Cubes[0].Perspectives)
            {
                //Create the Perspective
                TOM.Perspective TOMPerspective = new TOM.Perspective
                {
                    Name = AMOPerspective.Name,
                    Description = AMOPerspective.Description
                };
                foreach(AMO.PerspectiveDimension AMOPerspectiveDimension in AMOPerspective.Dimensions)
                {
                    TOM.Table TOMTable = TOMDatabase.Model.Tables[AMOPerspectiveDimension.Dimension.Name];
                    //Create the Perspective Table
                    TOM.PerspectiveTable TOMPerspectiveTable = new TOM.PerspectiveTable { Table = TOMTable };

                    //Add Columns
                    foreach (AMO.PerspectiveAttribute PerspectiveAttribute in AMOPerspectiveDimension.Attributes)
                        TOMPerspectiveTable.PerspectiveColumns.Add(new TOM.PerspectiveColumn
                        {
                            Column = TOMTable.Columns[PerspectiveAttribute.Attribute.Name]
                        });

                    //Add Hierarchies
                    foreach (AMO.PerspectiveHierarchy PerspectiveHierarchy in AMOPerspectiveDimension.Hierarchies)
                        TOMPerspectiveTable.PerspectiveHierarchies.Add(new TOM.PerspectiveHierarchy
                        {
                            Hierarchy = TOMTable.Hierarchies[PerspectiveHierarchy.Hierarchy.Name]
                        });

                    //Add Measures
                    foreach (AMO.PerspectiveCalculation PerspectiveCalculation in AMOPerspective.Calculations)
                    {
                        foreach(TOM.Measure TOMMeasure in TOMTable.Measures)
                            if(
                                PerspectiveCalculation.Type == AMO.PerspectiveCalculationType.Member 
                                && PerspectiveCalculation.Name == "[Measures].[" + TOMMeasure.Name + "]"
                            )
                                TOMPerspectiveTable.PerspectiveMeasures.Add(new TOM.PerspectiveMeasure{Measure = TOMMeasure});
                    }

                    //Add the Perspective Table
                    TOMPerspective.PerspectiveTables.Add(TOMPerspectiveTable);
                }
                //Add the Perspective
                TOMDatabase.Model.Perspectives.Add(TOMPerspective);
            }
            #endregion

            //Add TabularEditor compatibility annotation if necessary
            if(AddTabularEditorAnnotation)
                TOMDatabase.Model.Annotations.Add(new TOM.Annotation {
                    Name = "TabularEditor_CompatibilityVersion",
                    Value = AMODatabase.CompatibilityLevel.ToString()
                });
            //TODO: Handle KPIs
            return TOMDatabase;
        }
    }
}
