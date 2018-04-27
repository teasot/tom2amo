using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AMO = Microsoft.AnalysisServices;
using TOM = Microsoft.AnalysisServices.Tabular;
namespace TOM2AMO
{
    /* Based on AMO2Tabular (https://github.com/juanpablojofre/tabularamo/), which for all its flaws, was the first attempt at finding 
     * a way to programatically create tabular models. 
     * 
     * Without it, navgating th labyrinthine logic of the 1103 model would have been impossible.
     * So, shout outs to JuanPablo Jofre for creating the library which started it all!
     */
    public static class ToAMO
    {
        const int SQL2012RTM = 1100;
        const int SQL2012SP1 = 1103;
        const string MDXScriptName = "MdxScript";
        //TODO: Add direct query support
        /* Complete feature list, to be noted when complete support added
         *  - Database
         *      - Direct Query
         *      - Datasources (Done)
         *      - Tables (Done)
         *          - Translation of table (Done)
         *          - Attributes (Done)
         *              - Translation of Attribute (Done)
         *          - Hierarchies (Done)
         *              - Translation of Hierarchies (Done)
         *              - Levels (Done)
         *                  - Translation of Levels (Done)
         *      - Measures (Done)
         *          - Translation of Measures (Done)
         *          - KPI's (Done)
         *      - Perspectives (Done)
         *      - Roles (Done)
         *          - Row Level Security (Done)
         *          - Members (Done)
         *      - Relationships (Done)
         */
        /// <summary>
        /// Given a 1200 or 1400 model, produces an equivelant 1103 model. Incompatible features are not added.
        /// </summary>
        /// <param name="TOMDatabase"></param>
        /// <returns></returns>
        public static AMO.Database ToAMODatabase(TOM.Database TOMDatabase)
        {
            /* The Database representing a Tabular model rarely has a 1 to 1 mapping between
             * a TOM object and an AMO object.
             * 
             * Some objects which DO have 1:1 mappings (at least, in a logical sense):
             *  - DataSources
             *  - Hierarchies
             *  - Levels
             *  - Perspectives
             *  - Partitions
             *  - Translations (however, they are structured differently)
             *  - Annotations (These can also accept XML nodes, not just strings)
             *  
             * The following do NOT:
             *  - Tables
             *  - Columns
             *  - Measures
             *  - Relationships
             *  - Roles
             *  
             *  The following is, surprisingly, supported:
             *  - Tabular Actions
             *  - HideMemberIf
             *  - Display Folders
             *  - Custom Format strings
             *  - Translations
             *  - Similar to 1200, more than one column with the same source column
             *  
             * Please note this does NOT produce a file able to be opened by Visual Studio.
             * Unlike 1200 models, Visual Studio cannot open all valid 1103 models.
             * It relies heavily on custom annotations, understood mainly through trial and error.
             * 
             * In addition to these custom annotations, some features have no alternative (custom format strings), and
             * are completely unsupported by both Visual Studio and BIDS.
             * 
             * Display Folders, Tabular Actions, and Translations are all supported using BIDS.
             *  
             * These drastically increase the database/file size, and as such are optionally added in a seperate function.
             */

            TOM.Model TOMModel = TOMDatabase.Model;

            //Create Database
            AMO.Database AMODatabase = new AMO.Database(TOMDatabase.Name, TOMDatabase.Name);
            //Initialise with default values.
            AMODatabase.StorageEngineUsed = AMO.StorageEngineUsed.InMemory;
            AMODatabase.CompatibilityLevel = SQL2012SP1;

            //DataSource has 1:1 mapping with AMO object
            #region DataSources
            foreach (TOM.ProviderDataSource TOMDataSource in TOMModel.DataSources)
            {
                AMO.DataSource AMODataSource = new AMO.RelationalDataSource(TOMDataSource.Name, TOMDataSource.Name);
                AMODataSource.Description = TOMDataSource.Description;
                AMODataSource.ConnectionString = TOMDataSource.ConnectionString;
                AMODataSource.ImpersonationInfo = new AMO.ImpersonationInfo();
                switch (TOMDataSource.ImpersonationMode)
                {
                    case TOM.ImpersonationMode.Default:
                        AMODataSource.ImpersonationInfo.ImpersonationMode = AMO.ImpersonationMode.Default;
                        break;
                    case TOM.ImpersonationMode.ImpersonateAccount:
                        AMODataSource.ImpersonationInfo.ImpersonationMode = AMO.ImpersonationMode.ImpersonateAccount;
                        break;
                    case TOM.ImpersonationMode.ImpersonateAnonymous:
                        AMODataSource.ImpersonationInfo.ImpersonationMode = AMO.ImpersonationMode.ImpersonateAnonymous;
                        break;
                    case TOM.ImpersonationMode.ImpersonateCurrentUser:
                        AMODataSource.ImpersonationInfo.ImpersonationMode = AMO.ImpersonationMode.ImpersonateCurrentUser;
                        break;
                    case TOM.ImpersonationMode.ImpersonateServiceAccount:
                        AMODataSource.ImpersonationInfo.ImpersonationMode = AMO.ImpersonationMode.ImpersonateServiceAccount;
                        break;
                    case TOM.ImpersonationMode.ImpersonateUnattendedAccount:
                        AMODataSource.ImpersonationInfo.ImpersonationMode = AMO.ImpersonationMode.ImpersonateUnattendedAccount;
                        break;
                }
                AMODataSource.ImpersonationInfo.Account = TOMDataSource.Account;
                AMODataSource.ImpersonationInfo.Password = TOMDataSource.Password;
                switch (TOMDataSource.Isolation)
                {
                    case TOM.DatasourceIsolation.ReadCommitted:
                        AMODataSource.Isolation = AMO.DataSourceIsolation.ReadCommitted;
                        break;
                    case TOM.DatasourceIsolation.Snapshot:
                        AMODataSource.Isolation = AMO.DataSourceIsolation.Snapshot;
                        break;
                }
                //1 "tick" = 100 nanoseconds = 1*10^-7 seconds.
                AMODataSource.Timeout = new TimeSpan(TOMDataSource.Timeout * 10000000);
                AMODatabase.DataSources.Add(AMODataSource);
            }
            #endregion
            
            /* The DSV is surprisingly simple. 
             * For each physical table (but NOT partition), a DataTable needs to be added to the DSV DataSet,
             * with the same name as the table.
             * 
             * Similarly, for each distinct source column, a DataColumn needs to be added to the corresponding DataTable
             */
            #region DataSourceView
            using (AMO.DataSourceView dsv = new AMO.DataSourceView(AMODatabase.DataSources[0].Name))
            {
                System.Data.DataSet Schema = new System.Data.DataSet(AMODatabase.DataSources[0].Name);
                dsv.Schema = Schema;
                dsv.DataSourceID = AMODatabase.DataSources[0].Name;
                AMODatabase.DataSourceViews.Add(dsv);
            }
            #endregion

            #region Create Cube
            AMO.Cube Cube = new AMO.Cube(TOMModel.Name, TOMModel.Name);
            
            Cube.Source = new AMO.DataSourceViewBinding(AMODatabase.DataSourceViews[0].ID);
            Cube.StorageMode = AMO.StorageMode.InMemory;

            //Create the MdxScript for holding DAX commands.
            using (AMO.MdxScript Script = Cube.MdxScripts.Add(MDXScriptName, MDXScriptName))
            {
                //You MUST have a "default" MDX measure for some reason. 
                //We make sure to hide it - it serves no real purpose, besides its own existence
                System.Text.StringBuilder InitialisationCommand = new System.Text.StringBuilder();
                InitialisationCommand.AppendLine("CALCULATE;");
                InitialisationCommand.AppendLine("CREATE MEMBER CURRENTCUBE.Measures.[_No measures defined] AS 1, VISIBLE = 0;");
                InitialisationCommand.AppendLine("ALTER CUBE CURRENTCUBE UPDATE DIMENSION Measures, Default_Member = [_No measures defined];");
                Script.Commands.Add(new AMO.Command(InitialisationCommand.ToString()));
            }
            AMODatabase.Cubes.Add(Cube);
            #endregion
            #region Add Tables
            foreach (TOM.Table TOMTable in TOMModel.Tables)
            {
                //Three "parts" to a Table:
                //1. A System.Data.DataTable, with the same Name as the TOMTable
                //2. A dimension with the same name as the TOMTable
                //3. A measure group with the same name as the TOMTable, 
                System.Data.DataTable SchemaTable = new System.Data.DataTable(TOMTable.Name);
                AMODatabase.DataSourceViews[0].Schema.Tables.Add(SchemaTable);

                string RowNumberColumnName = string.Format(System.Globalization.CultureInfo.InvariantCulture, "RowNumber_{0}", System.Guid.NewGuid());

                #region Add Table Dimension
                try
                {
                    using (AMO.Dimension Dimension = AMODatabase.Dimensions.Add(TOMTable.Name, TOMTable.Name))
                    {
                        Dimension.Source = new AMO.DataSourceViewBinding(AMODatabase.DataSourceViews[0].ID);
                        Dimension.StorageMode = AMO.DimensionStorageMode.InMemory;
                        Dimension.UnknownMember = AMO.UnknownMemberBehavior.AutomaticNull;
                        Dimension.UnknownMemberName = "Unknown";
                        using (Dimension.ErrorConfiguration = new AMO.ErrorConfiguration())
                        {
                            Dimension.ErrorConfiguration.KeyNotFound = AMO.ErrorOption.IgnoreError;
                            Dimension.ErrorConfiguration.KeyDuplicate = AMO.ErrorOption.ReportAndStop;
                            Dimension.ErrorConfiguration.NullKeyNotAllowed = AMO.ErrorOption.ReportAndStop;
                        }
                        Dimension.ProactiveCaching = new AMO.ProactiveCaching();
                        System.TimeSpan DefaultProactiveCachingTimeSpan = new System.TimeSpan(0, 0, -1);
                        Dimension.ProactiveCaching.SilenceInterval = DefaultProactiveCachingTimeSpan;
                        Dimension.ProactiveCaching.Latency = DefaultProactiveCachingTimeSpan;
                        Dimension.ProactiveCaching.SilenceOverrideInterval = DefaultProactiveCachingTimeSpan;
                        Dimension.ProactiveCaching.ForceRebuildInterval = DefaultProactiveCachingTimeSpan;
                        Dimension.ProactiveCaching.Source = new AMO.ProactiveCachingInheritedBinding();
                        
                        Dimension.Description = TOMTable.Description;

                        // Define RowNumber
                        using (AMO.DimensionAttribute RowNumberDimAttribute = Dimension.Attributes.Add(RowNumberColumnName, RowNumberColumnName))
                        {
                            RowNumberDimAttribute.Type = AMO.AttributeType.RowNumber;
                            RowNumberDimAttribute.KeyUniquenessGuarantee = true;
                            RowNumberDimAttribute.Usage = AMO.AttributeUsage.Key;
                            RowNumberDimAttribute.KeyColumns.Add(new AMO.DataItem());
                            RowNumberDimAttribute.KeyColumns[0].DataType = System.Data.OleDb.OleDbType.Integer;
                            RowNumberDimAttribute.KeyColumns[0].DataSize = 4;
                            RowNumberDimAttribute.KeyColumns[0].NullProcessing = AMO.NullProcessing.Error;
                            RowNumberDimAttribute.KeyColumns[0].Source = new AMO.RowNumberBinding();
                            RowNumberDimAttribute.NameColumn = new AMO.DataItem();
                            RowNumberDimAttribute.NameColumn.DataType = System.Data.OleDb.OleDbType.WChar;
                            RowNumberDimAttribute.NameColumn.DataSize = 4;
                            RowNumberDimAttribute.NameColumn.NullProcessing = AMO.NullProcessing.ZeroOrBlank;
                            RowNumberDimAttribute.NameColumn.Source = new AMO.RowNumberBinding();
                            RowNumberDimAttribute.OrderBy = AMO.OrderBy.Key;
                            RowNumberDimAttribute.AttributeHierarchyVisible = false;
                        }

                        // Add Translations
                        foreach (TOM.Culture TOMCulture in TOMModel.Cultures)
                        {
                            AMO.Translation AMOTranslation = TranslationHelper.GetTranslation(TOMCulture, TOMTable);
                            if(AMOTranslation != null)
                                Dimension.Translations.Add(AMOTranslation);
                        }
                    }
                }
                catch (Exception e)
                {
                    throw new Exception(string.Format("The following error occurred creating the DImension for table {0}: {1}", TOMTable.Name, e.Message), e);
                }
                #endregion
                #region Add Table MeasureGroup
                try
                {
                    using (AMO.MeasureGroup TableMeasureGroup = Cube.MeasureGroups.Add(TOMTable.Name, TOMTable.Name))
                    {
                        TableMeasureGroup.StorageMode = AMO.StorageMode.InMemory;
                        TableMeasureGroup.ProcessingMode = AMO.ProcessingMode.Regular;

                        // Add Default Measure
                        string DefaultMeasureID = string.Concat("_Count", TOMTable.Name);
                        using (AMO.Measure DefaultMeasure = TableMeasureGroup.Measures.Add(DefaultMeasureID, DefaultMeasureID))
                        using (AMO.RowBinding DefaultMeasureRowBinding = new AMO.RowBinding(TOMTable.Name))
                        using (AMO.DataItem DefaultMeasureSource = new AMO.DataItem(DefaultMeasureRowBinding))
                        {
                            DefaultMeasure.AggregateFunction = AMO.AggregationFunction.Count;
                            DefaultMeasure.DataType = AMO.MeasureDataType.BigInt;
                            DefaultMeasure.Visible = false;
                            DefaultMeasureSource.DataType = System.Data.OleDb.OleDbType.BigInt;
                            DefaultMeasure.Source = DefaultMeasureSource;
                        }

                        // Add Dimension to Measure Group
                        using (AMO.DegenerateMeasureGroupDimension DefaultMeasureGroupDimension = new AMO.DegenerateMeasureGroupDimension(TOMTable.Name))
                        using (AMO.MeasureGroupAttribute MeasureGroupAttribute = new AMO.MeasureGroupAttribute(RowNumberColumnName))
                        using (AMO.ColumnBinding RowNumberColumnBinding = new AMO.ColumnBinding(TOMTable.Name, RowNumberColumnName))
                        using (AMO.DataItem RowNumberKeyColumn = new AMO.DataItem(RowNumberColumnBinding))
                        {
                            DefaultMeasureGroupDimension.ShareDimensionStorage = AMO.StorageSharingMode.Shared;
                            DefaultMeasureGroupDimension.CubeDimensionID = TOMTable.Name;
                            MeasureGroupAttribute.Type = AMO.MeasureGroupAttributeType.Granularity;
                            RowNumberKeyColumn.DataType = System.Data.OleDb.OleDbType.Integer;
                            MeasureGroupAttribute.KeyColumns.Add(RowNumberKeyColumn);
                            DefaultMeasureGroupDimension.Attributes.Add(MeasureGroupAttribute);
                            TableMeasureGroup.Dimensions.Add(DefaultMeasureGroupDimension);
                        }

                        //Partitions have a 1:1 mapping
                        #region Partitions
                        foreach (TOM.Partition TOMPartition in TOMTable.Partitions)
                        {
                            using (AMO.Partition AMOPartition = new AMO.Partition(TOMPartition.Name, TOMPartition.Name))
                            {
                                AMOPartition.StorageMode = AMO.StorageMode.InMemory;
                                AMOPartition.ProcessingMode = AMO.ProcessingMode.Regular;
                                AMOPartition.Source = new AMO.QueryBinding(
                                    ((TOM.QueryPartitionSource)TOMPartition.Source).DataSource.Name, 
                                    ((TOM.QueryPartitionSource)TOMPartition.Source).Query
                                );
                                AMOPartition.Type = AMO.PartitionType.Data;
                                TableMeasureGroup.Partitions.Add(AMOPartition);
                            }
                        }
                        #endregion
                    }
                }
                catch (Exception e)
                {
                    throw new Exception(string.Format("The following error occurred creating the MeasureGroup for table {0}: {1}", TOMTable.Name, e.Message), e);
                }
                #endregion

                //Add the dimension to the cube
                Cube.Dimensions.Add(TOMTable.Name, TOMTable.Name, TOMTable.Name);
                Cube.Dimensions[TOMTable.Name].Visible = !TOMTable.IsHidden;

                /* No 1:1 mapping exists for columns. Each column consists of the following:
                 *  - A DataColumn with the same name as the column,
                 *      attached to the DataTable correlating to the Table, located in the DSV
                 *  - An attribute with the same name as the column, added to the Dimension
                 *  - An AttributeRelationship with the RowNumber column
                 */

                #region Add Columns
                foreach (TOM.Column TOMColumn in TOMTable.Columns)
                {

                    switch (TOMColumn.Type)
                    {
                        case TOM.ColumnType.Data:
                            //Add the DataColumn corresponding to the SourceColumn, if it does not already exist
                            TOM.DataColumn TOMDataColumn = (TOM.DataColumn)TOMColumn;
                            if (!SchemaTable.Columns.Contains(TOMDataColumn.SourceColumn))
                                SchemaTable.Columns.Add(new System.Data.DataColumn(((TOM.DataColumn)TOMColumn).SourceColumn));

                            System.Data.OleDb.OleDbType ColumnDataType = DataTypeHelper.ToOleDbType(TOMDataColumn.DataType);
                            AMO.DimensionAttribute NormalAttribute = AMODatabase.Dimensions[TOMTable.Name].Attributes.Add(TOMDataColumn.Name, TOMDataColumn.Name);
                            NormalAttribute.Usage = AMO.AttributeUsage.Regular;
                            NormalAttribute.KeyUniquenessGuarantee = false;
                            NormalAttribute.KeyColumns.Add(new AMO.DataItem(SchemaTable.TableName, SchemaTable.Columns[TOMDataColumn.SourceColumn].ColumnName, ColumnDataType));
                            NormalAttribute.KeyColumns[0].Source = new AMO.ColumnBinding(SchemaTable.TableName, SchemaTable.Columns[TOMDataColumn.SourceColumn].ColumnName);
                            NormalAttribute.KeyColumns[0].NullProcessing = AMO.NullProcessing.Preserve;
                            NormalAttribute.NameColumn = new AMO.DataItem(SchemaTable.TableName, SchemaTable.Columns[TOMDataColumn.SourceColumn].ColumnName, System.Data.OleDb.OleDbType.WChar);
                            NormalAttribute.NameColumn.Source = new AMO.ColumnBinding(SchemaTable.TableName, SchemaTable.Columns[TOMDataColumn.SourceColumn].ColumnName);
                            NormalAttribute.NameColumn.NullProcessing = AMO.NullProcessing.ZeroOrBlank;
                            NormalAttribute.OrderBy = AMO.OrderBy.Key;
                            AMO.AttributeRelationship NormalAttributeRelationship = AMODatabase.Dimensions[TOMTable.Name].Attributes[RowNumberColumnName].AttributeRelationships.Add(NormalAttribute.ID);

                            NormalAttribute.AttributeHierarchyVisible = !TOMDataColumn.IsHidden;
                            NormalAttribute.Description = TOMDataColumn.Description;
                            NormalAttribute.AttributeHierarchyDisplayFolder = TOMDataColumn.DisplayFolder;
                            //Add Translations to the CalculatedAttribute
                            foreach (TOM.Culture TOMCulture in TOMModel.Cultures)
                                using (AMO.Translation AMOTranslation = TranslationHelper.GetTranslation(TOMCulture, TOMColumn))
                                    if (AMOTranslation != null)
                                        NormalAttribute.Translations.Add(new AMO.AttributeTranslation { Caption = AMOTranslation.Caption, Description = AMOTranslation.Description, DisplayFolder = AMOTranslation.DisplayFolder });

                            NormalAttributeRelationship.Cardinality = AMO.Cardinality.Many;
                            NormalAttributeRelationship.OverrideBehavior = AMO.OverrideBehavior.None;
                            break;
                        case TOM.ColumnType.Calculated:
                            TOM.CalculatedColumn TOMCalculatedColumn = (TOM.CalculatedColumn)TOMColumn;
                            System.Data.OleDb.OleDbType CalculatedColumnDataType = DataTypeHelper.ToOleDbType(TOMColumn.DataType);

                            //Add Attribute to the Dimension
                            AMO.Dimension dim = AMODatabase.Dimensions[TOMTable.Name];
                            AMO.DimensionAttribute CalculatedAttribute = dim.Attributes.Add(TOMCalculatedColumn.Name, TOMCalculatedColumn.Name);
                            CalculatedAttribute.Usage = AMO.AttributeUsage.Regular;
                            CalculatedAttribute.KeyUniquenessGuarantee = false;

                            CalculatedAttribute.KeyColumns.Add(new AMO.DataItem(TOMTable.Name, TOMCalculatedColumn.Name, CalculatedColumnDataType));
                            CalculatedAttribute.KeyColumns[0].Source = new AMO.ExpressionBinding(TOMCalculatedColumn.Expression);
                            CalculatedAttribute.KeyColumns[0].NullProcessing = AMO.NullProcessing.Preserve;
                            CalculatedAttribute.NameColumn = new AMO.DataItem(TOMTable.Name, TOMCalculatedColumn.Name, System.Data.OleDb.OleDbType.WChar);
                            CalculatedAttribute.NameColumn.Source = new AMO.ExpressionBinding(TOMCalculatedColumn.Expression);
                            CalculatedAttribute.NameColumn.NullProcessing = AMO.NullProcessing.ZeroOrBlank;

                            CalculatedAttribute.OrderBy = AMO.OrderBy.Key;
                            AMO.AttributeRelationship currentAttributeRelationship = dim.Attributes[RowNumberColumnName].AttributeRelationships.Add(CalculatedAttribute.ID);

                            CalculatedAttribute.AttributeHierarchyVisible = !TOMCalculatedColumn.IsHidden;
                            CalculatedAttribute.Description = TOMCalculatedColumn.Description;
                            CalculatedAttribute.AttributeHierarchyDisplayFolder = TOMCalculatedColumn.DisplayFolder;
                            //Add Translations to the CalculatedAttribute
                            //Loop through each culture, and add the translation associated with that culture.
                            foreach (TOM.Culture TOMCulture in TOMModel.Cultures)
                                using (AMO.Translation AMOTranslation = TranslationHelper.GetTranslation(TOMCulture, TOMColumn))
                                    if (AMOTranslation != null)
                                        CalculatedAttribute.Translations.Add(new AMO.AttributeTranslation { Caption = AMOTranslation.Caption, Description = AMOTranslation.Description, DisplayFolder = AMOTranslation.DisplayFolder });

                            currentAttributeRelationship.Cardinality = AMO.Cardinality.Many;
                            currentAttributeRelationship.OverrideBehavior = AMO.OverrideBehavior.None;

                            //Add CalculatedColumn as attribute to the MeasureGroup
                            AMO.MeasureGroup mg = Cube.MeasureGroups[TOMTable.Name];
                            AMO.DegenerateMeasureGroupDimension currentMGDim = (AMO.DegenerateMeasureGroupDimension)mg.Dimensions[TOMTable.Name];
                            AMO.MeasureGroupAttribute mga = new AMO.MeasureGroupAttribute(TOMCalculatedColumn.Name);

                            mga.KeyColumns.Add(new AMO.DataItem(TOMTable.Name, TOMCalculatedColumn.Name, System.Data.OleDb.OleDbType.Empty));
                            mga.KeyColumns[0].Source = new AMO.ExpressionBinding(TOMCalculatedColumn.Expression);
                            currentMGDim.Attributes.Add(mga);
                            break;
                        default:
                            throw new System.NotImplementedException(string.Format("Cannot deploy Column of type {0}", TOMColumn.Type.ToString()));
                    }
                }
                #endregion

                //Add sort by columns
                foreach (TOM.Column TOMColumn in TOMTable.Columns)
                    if (TOMColumn.SortByColumn != null)
                        AMODatabase.Dimensions[TOMTable.Name].Attributes[TOMColumn.Name].OrderByAttributeID = TOMColumn.SortByColumn.Name;

                #region Add Hierarchies
                foreach (TOM.Hierarchy TOMHierarchy in TOMTable.Hierarchies)
                {
                    //Create the Hierarchy, and add it
                    AMO.Hierarchy AMOHierarchy = AMODatabase.Dimensions[TOMTable.Name].Hierarchies.Add(TOMHierarchy.Name, TOMHierarchy.Name);
                    AMOHierarchy.Description = TOMHierarchy.Description;
                    AMOHierarchy.DisplayFolder = TOMHierarchy.DisplayFolder;

                    AMOHierarchy.AllMemberName = "All";
                    foreach (TOM.Level TOMLevel in TOMHierarchy.Levels)
                    {
                        AMO.Level AMOLevel = AMOHierarchy.Levels.Add(TOMLevel.Name);
                        AMOLevel.SourceAttribute = AMODatabase.Dimensions[TOMTable.Name].Attributes[TOMLevel.Column.Name];
                        AMOLevel.Description = TOMLevel.Description;
                        //Add Translations to the CalculatedAttribute
                        //Loop through each culture, and add the translation associated with that culture.
                        foreach (TOM.Culture TOMCulture in TOMModel.Cultures)
                            using (AMO.Translation AMOTranslation = TranslationHelper.GetTranslation(TOMCulture, TOMLevel))
                                if (AMOTranslation != null)
                                    AMOLevel.Translations.Add(AMOTranslation);
                    }
                }
                #endregion
            }
            #endregion

            #region Add Measures
            using (AMO.MdxScript MdxScript = Cube.DefaultMdxScript)
            {
                //Create a "default" measure, required for the cube to function.
                MdxScript.Commands.Add(new AMO.Command("CALCULATE;"
                + "CREATE MEMBER CURRENTCUBE.Measures.[_No measures defined] AS 1, VISIBLE = 0;"
                + "ALTER CUBE CURRENTCUBE UPDATE DIMENSION Measures, Default_Member = [_No measures defined];"));

                foreach (TOM.Table TOMTable in TOMModel.Tables)
                {
                    foreach (TOM.Measure TOMMeasure in TOMTable.Measures)
                    {
                        //Create the Command, which contains the definition of the Measure
                        AMO.Command AMOCommand = new AMO.Command(string.Format("CREATE MEASURE '{0}'[{1}]={2};", TOMTable.Name, TOMMeasure.Name.Replace("]", "]]"), TOMMeasure.Expression));
                        if (TOMMeasure.KPI != null)
                        {
                            #region Add KPI
                            //Start building the final command - we add to it as we go along
                            string FinalKPICreate = string.Format(
                                "CREATE KPI CURRENTCUBE.[{1}] AS Measures.[{1}], ASSOCIATED_MEASURE_GROUP = '{0}'",
                                TOMTable.Name,
                                TOMMeasure.Name
                            );
                            //Goal/Target
                            if (!string.IsNullOrWhiteSpace(TOMMeasure.KPI.TargetExpression)) {
                                AMOCommand.Text += Environment.NewLine + string.Format(string.Format("CREATE MEASURE '{0}'[_{1} Goal]={2};", TOMTable.Name, TOMMeasure.Name.Replace("]", "]]"), TOMMeasure.KPI.TargetExpression));
                                FinalKPICreate += string.Format(", GOAL = Measures.[_{0} Goal]", TOMMeasure.Name.Replace("]", "]]"));
                            }

                            //Status
                            if (!string.IsNullOrWhiteSpace(TOMMeasure.KPI.StatusExpression)) {
                                AMOCommand.Text += Environment.NewLine + string.Format(string.Format("CREATE MEASURE '{0}'[_{1} Status]={2};", TOMTable.Name, TOMMeasure.Name.Replace("]", "]]"), TOMMeasure.KPI.StatusExpression));
                                FinalKPICreate += string.Format(", STATUS = Measures.[_{0} Status]", TOMMeasure.Name.Replace("]", "]]"));
                                if (!string.IsNullOrWhiteSpace(TOMMeasure.KPI.StatusGraphic))
                                    FinalKPICreate += string.Format(", STATUS_GRAPHIC = '{0}'", TOMMeasure.KPI.StatusGraphic);
                            }
                            
                            //Trend
                            if (!string.IsNullOrWhiteSpace(TOMMeasure.KPI.TrendExpression)) {
                                AMOCommand.Text += Environment.NewLine + string.Format(string.Format("CREATE MEASURE '{0}'[_{1} Trend]={2};", TOMTable.Name, TOMMeasure.Name.Replace("]", "]]"), TOMMeasure.KPI.TrendExpression));
                                FinalKPICreate += string.Format(", TREND = Measures.[_{0} Trend]", TOMMeasure.Name.Replace("]", "]]"));
                                if (!string.IsNullOrWhiteSpace(TOMMeasure.KPI.TrendGraphic))
                                    FinalKPICreate += string.Format(", TREND_GRAPHIC = '{0}'", TOMMeasure.KPI.TrendGraphic);
                            }

                            AMOCommand.Text += Environment.NewLine + FinalKPICreate + ";";

                            //Create calculation properties, hiding the "fake" measures (if they exist)
                            //Target
                            if (!string.IsNullOrWhiteSpace(TOMMeasure.KPI.TargetExpression))
                            {
                                AMO.CalculationProperty TargetCalculationProperty = new AMO.CalculationProperty(string.Format("[_{0} Goal]", TOMMeasure.Name.Replace("]", "]]")), AMO.CalculationType.Member);
                                TargetCalculationProperty.Description = TOMMeasure.KPI.TargetDescription;
                                TargetCalculationProperty.CalculationType = AMO.CalculationType.Member;
                                TargetCalculationProperty.Visible = false;
                                if (!string.IsNullOrWhiteSpace(TOMMeasure.KPI.TargetFormatString))
                                    TargetCalculationProperty.FormatString = "'" + TOMMeasure.KPI.TargetFormatString + "'";
                                MdxScript.CalculationProperties.Add(TargetCalculationProperty);
                            }

                            //Status
                            if (!string.IsNullOrWhiteSpace(TOMMeasure.KPI.StatusExpression))
                            {
                                AMO.CalculationProperty StatusCalculationProperty = new AMO.CalculationProperty(string.Format("[_{0} Status]", TOMMeasure.Name.Replace("]", "]]")), AMO.CalculationType.Member);
                                StatusCalculationProperty.Description = TOMMeasure.KPI.StatusDescription;
                                StatusCalculationProperty.CalculationType = AMO.CalculationType.Member;
                                StatusCalculationProperty.Visible = false;
                                MdxScript.CalculationProperties.Add(StatusCalculationProperty);
                            }

                            //Trend
                            if (!string.IsNullOrWhiteSpace(TOMMeasure.KPI.StatusExpression))
                            {
                                AMO.CalculationProperty TrendCalculationProperty = new AMO.CalculationProperty(string.Format("[_{0} Trend]", TOMMeasure.Name.Replace("]", "]]")), AMO.CalculationType.Member);
                                TrendCalculationProperty.Description = TOMMeasure.KPI.TrendDescription;
                                TrendCalculationProperty.CalculationType = AMO.CalculationType.Member;
                                TrendCalculationProperty.Visible = false;
                                MdxScript.CalculationProperties.Add(TrendCalculationProperty);
                            }

                            //Create the KPI calculation property
                            AMO.CalculationProperty KPICalculationProperty = new AMO.CalculationProperty(string.Format("KPIs.[{0}]", TOMMeasure.Name.Replace("]", "]]")), AMO.CalculationType.Member);
                            KPICalculationProperty.Description = TOMMeasure.KPI.Description;
                            KPICalculationProperty.CalculationType = AMO.CalculationType.Member;
                            MdxScript.CalculationProperties.Add(KPICalculationProperty);
                            #endregion
                        }

                        //Add the Command to the MdxScript
                        MdxScript.Commands.Add(AMOCommand);

                        //Create the Calculation Property, which contains the various properties of the Measure
                        AMO.CalculationProperty CalculationProperty = new AMO.CalculationProperty(string.Format("[{0}]", TOMMeasure.Name.Replace("]", "]]")), AMO.CalculationType.Member);
                        CalculationProperty.Description = TOMMeasure.Description;
                        CalculationProperty.DisplayFolder = TOMMeasure.DisplayFolder;
                        CalculationProperty.CalculationType = AMO.CalculationType.Member;
                        CalculationProperty.Visible = !TOMMeasure.IsHidden;
                        if (!string.IsNullOrWhiteSpace(TOMMeasure.FormatString))
                            CalculationProperty.FormatString = "'" + TOMMeasure.FormatString + "'";

                        //Add Translations to the Calculation property
                        foreach (TOM.Culture TOMCulture in TOMModel.Cultures)
                            using (AMO.Translation AMOTranslation = TranslationHelper.GetTranslation(TOMCulture, TOMMeasure))
                                if (AMOTranslation != null)
                                    CalculationProperty.Translations.Add(AMOTranslation);

                        //Finally, add the CalculationProperty to the MDX script
                        MdxScript.CalculationProperties.Add(CalculationProperty);

                    }
                }
            }
            #endregion

            #region Add Relationships
            //Relationships are just TOO awful. They are placed in their own helper function as a result.
            foreach (TOM.SingleColumnRelationship TOMRelationship in TOMModel.Relationships)
                RelationshipHelper.CreateRelationship(TOMRelationship, AMODatabase);
            #endregion

            #region Add Perspectives
            foreach(TOM.Perspective TOMPerspective in TOMModel.Perspectives)
            {
                AMO.Perspective AMOPerspective = new AMO.Perspective(TOMPerspective.Name);
                foreach(TOM.PerspectiveTable TOMTable in TOMPerspective.PerspectiveTables)
                {
                    //Add Perspective Dimension
                    AMO.PerspectiveDimension PerspectiveDimension = new AMO.PerspectiveDimension(TOMTable.Name);

                    //Perspective Columns
                    foreach (TOM.PerspectiveColumn TOMPerspectiveColumn in TOMTable.PerspectiveColumns)
                        PerspectiveDimension.Attributes.Add(TOMPerspectiveColumn.Name);
                    //Perspective Hierarchies
                    foreach (TOM.PerspectiveHierarchy TOMPerspectiveHierarchy in TOMTable.PerspectiveHierarchies)
                        PerspectiveDimension.Hierarchies.Add(TOMPerspectiveHierarchy.Name);

                    //Add Perspective MeasureGroup
                    AMO.PerspectiveMeasureGroup PerspectiveMeasureGroup = new AMO.PerspectiveMeasureGroup(TOMTable.Name);

                    //Perspective Measures
                    //In this case, ']' is NOT "double quoted", unlike the calculation references, s
                    foreach (TOM.PerspectiveMeasure TOMPerspectiveMeasure in TOMTable.PerspectiveMeasures)
                        AMOPerspective.Calculations.Add('[' + TOMPerspectiveMeasure.Name + ']');

                    //VS does not add KPI's to Perspectives, so I have no idea how to do this...
                    //TODO: Add KPIs to perspectives
                    
                }
                AMODatabase.Cubes[0].Perspectives.Add(AMOPerspective);
            }
            #endregion

            #region Roles
            foreach (TOM.ModelRole TOMRole in TOMDatabase.Model.Roles)
            {
                //Ceate database role
                AMO.Role DatabaseRole = AMODatabase.Roles.Add(TOMRole.Name);
                DatabaseRole.Description = TOMRole.Description;

                //Add Members to Database role
                foreach(TOM.ModelRoleMember TOMMember in TOMRole.Members)
                {
                    DatabaseRole.Members.Add(new AMO.RoleMember
                    {
                        Name = TOMMember.MemberName,
                        Sid = TOMMember.MemberID
                    });
                }

                //Add DatabasePermission
                AMO.DatabasePermission DatabasePermission = AMODatabase.DatabasePermissions.Add(TOMRole.Name, TOMRole.Name, TOMRole.Name);
                //Add CubePermission
                AMO.CubePermission CubePermission = AMODatabase.Cubes[0].CubePermissions.Add(TOMRole.Name, TOMRole.Name, TOMRole.Name);
                switch (TOMRole.ModelPermission)
                {
                    case TOM.ModelPermission.Administrator:
                        //Add ReadDefinition, Read, and Administer to DatabasePermission
                        DatabasePermission.ReadDefinition = AMO.ReadDefinitionAccess.Allowed;
                        DatabasePermission.Read = AMO.ReadAccess.Allowed;
                        DatabasePermission.Administer = true;

                        //Nothing extra to add to CubePermission, that is not added after this
                        // switch statement
                        break;
                    case TOM.ModelPermission.None:
                        //Add no permissions to DatabasePermission.
                        break;
                    case TOM.ModelPermission.Read:
                        //Add only Read access to Database Permission
                        DatabasePermission.Read = AMO.ReadAccess.Allowed;

                        //CubePermission
                        CubePermission.Read = AMO.ReadAccess.Allowed;
                        break;
                    case TOM.ModelPermission.ReadRefresh:
                        DatabasePermission.Read = AMO.ReadAccess.Allowed;
                        DatabasePermission.Process = true;

                        //CubePermission
                        CubePermission.Read = AMO.ReadAccess.Allowed;
                        CubePermission.Process = true;
                        break;
                    case TOM.ModelPermission.Refresh:
                        DatabasePermission.Process = true;

                        //CubePermission
                        CubePermission.Process = true;
                        break;
                }
                CubePermission.ReadSourceData = AMO.ReadSourceDataAccess.None;

                //Finally, add the row filters
                foreach(TOM.TablePermission TablePermission in TOMRole.TablePermissions)
                {
                    AMO.DimensionPermission DimensionPermission = new AMO.DimensionPermission(TOMRole.Name, TOMRole.Name, TOMRole.Name);
                    DimensionPermission.AllowedRowsExpression = TablePermission.FilterExpression;
                    AMODatabase.Dimensions.GetByName(TablePermission.Table.Name).DimensionPermissions.Add(DimensionPermission);
                }
            }
            #endregion
            return AMODatabase;
        }

        //TODO: Add
        /*TODO: Document necessary annotations/comments before addition.
         * Off the top of my head: 
         *   The "arbritrary" comment before measure declaration
         *   Query Editor Serialisation
         *   Annotation telling item to use serialisation, and not treat as table
         *   Data Table Annotations
         *   CalculationProperty annotations, including detecting common format strings
         *  Most definitely NOT BIDS - that is added in another function
         * 
         */
        public static void MakeCompatibleWithVisualStudio(AMO.Database AMODatabase)
        {
            throw new NotImplementedException("MakeCompatibleWithVisualStudio not yet implemented.");
        }

        public static void MakeCompatibleWithBIDS(AMO.Database AMODatabase)
        {
            throw new NotImplementedException("MakeCompatibleWithBIDS not yet implemented.");
        }
    }
}