﻿using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using TOM2AMO.Parsing;
using AMO = Microsoft.AnalysisServices;
using TOM = Microsoft.AnalysisServices.Tabular;
namespace TOM2AMO
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
         *      - Perspectives (Done)
         *      - Roles (Done)
         *          - Row Level Security (Done)
         *          - Members (Done)
         *      - Relationships (Done)
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
                        //Declare "generic" TOM Column, to be assigned to the "specific" column for reuse in
                        // assigning common properties
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
                        
                        //Generic Properties, shared between both Data Columns and Calculated Columns
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
                //Add sort by columns last
                //This is because we cannot add sort by columns referring to columns which do not exist yet
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
                    TOM.Partition TOMPartition = new TOM.Partition
                    {
                        Description = AMOPartition.Description,

                        //Add the query
                        Source = new TOM.QueryPartitionSource
                        {
                            DataSource = TOMDatabase.Model.DataSources[AMOPartition.DataSource.Name],
                            Query = ((AMO.QueryBinding)AMOPartition.Source).QueryDefinition
                        }
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
                        /* We do not care if a command is claid, only if we can parse it.
                         * We need two strings: 
                         * - A single quoted string, representing the table name
                         * - A square bracketed string, representing the measure name
                         * As long as they are present, we consider it valid.
                         */
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
                    //Find the TOM Table equivelant to the AMO Perspective Dimension
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
                        foreach (TOM.Measure TOMMeasure in TOMTable.Measures)
                            if (
                                PerspectiveCalculation.Type == AMO.PerspectiveCalculationType.Member
                                && PerspectiveCalculation.Name == "[Measures].[" + TOMMeasure.Name + "]"
                            )
                            {
                                TOMPerspectiveTable.PerspectiveMeasures.Add(new TOM.PerspectiveMeasure { Measure = TOMMeasure });
                            }
                    }

                    //Add the Perspective Table
                    TOMPerspective.PerspectiveTables.Add(TOMPerspectiveTable);
                }
                //Add the Perspective
                TOMDatabase.Model.Perspectives.Add(TOMPerspective);
            }
            #endregion

            #region Roles
            foreach(AMO.Role AMORole in AMODatabase.Roles)
            {
                //Create the Role
                TOM.ModelRole TOMRole = new TOM.ModelRole
                {
                    Name = AMORole.Name,
                    Description = AMORole.Description
                };

                //Determine the ModelPermission from the equivelant DatabasePermission
                foreach(AMO.DatabasePermission Permission in AMODatabase.DatabasePermissions)
                {
                    if(Permission.Role.Name == AMORole.Name)
                    {
                        if (Permission.Administer)
                            TOMRole.ModelPermission = TOM.ModelPermission.Administrator;
                        else if (Permission.Read == AMO.ReadAccess.Allowed && Permission.Process)
                            TOMRole.ModelPermission = TOM.ModelPermission.ReadRefresh;
                        else if (Permission.Read == AMO.ReadAccess.Allowed)
                            TOMRole.ModelPermission = TOM.ModelPermission.Read;
                        else if (Permission.Process)
                            TOMRole.ModelPermission = TOM.ModelPermission.Refresh;
                        else
                            TOMRole.ModelPermission = TOM.ModelPermission.None;
                    }
                }

                //Add the Row Level Security
                foreach(AMO.Dimension Dimension in AMODatabase.Dimensions)
                {
                    foreach (AMO.DimensionPermission DimensionPermission in Dimension.DimensionPermissions)
                        if (DimensionPermission.Role.Name == TOMRole.Name)
                        {
                            //Create the Table Permission
                            TOM.TablePermission TablePermission = new TOM.TablePermission
                            {
                                Table = TOMDatabase.Model.Tables[Dimension.Name],
                                FilterExpression = DimensionPermission.AllowedRowsExpression
                            };

                            //Add the Table Permission to the Role
                            TOMRole.TablePermissions.Add(TablePermission);
                        }
                }

                //Add Role Members
                foreach(AMO.RoleMember AMOMember in AMORole.Members)
                {
                    //Create the Role Member
                    TOM.ModelRoleMember TOMMember = new TOM.WindowsModelRoleMember
                    {
                        MemberID = AMOMember.Sid,
                        MemberName = AMOMember.Name
                    };

                    //Add the Member to the Role
                    TOMRole.Members.Add(TOMMember);
                }

                //Add the Role to the Database
                TOMDatabase.Model.Roles.Add(TOMRole);
            }
            #endregion

            #region Relationships
            foreach(AMO.Dimension Dimension in AMODatabase.Dimensions)
            {
                foreach(AMO.Relationship AMORelationship in Dimension.Relationships)
                {
                    //Get To and From columns
                    AMO.Dimension FromDimension = AMODatabase.Dimensions[AMORelationship.FromRelationshipEnd.DimensionID];
                    AMO.Dimension ToDimension = AMODatabase.Dimensions[AMORelationship.ToRelationshipEnd.DimensionID];
                    TOM.Table FromTable = TOMDatabase.Model.Tables[FromDimension.Name];
                    TOM.Column FromColumn = FromTable.Columns[FromDimension.Attributes[AMORelationship.FromRelationshipEnd.Attributes[0].AttributeID].Name];
                    TOM.Table ToTable = TOMDatabase.Model.Tables[ToDimension.Name];
                    TOM.Column ToColumn = ToTable.Columns[ToDimension.Attributes[AMORelationship.ToRelationshipEnd.Attributes[0].AttributeID].Name];

                    //Create the Relationship
                    TOM.SingleColumnRelationship TOMRelationship = new TOM.SingleColumnRelationship
                    {
                        FromColumn = FromColumn,
                        ToColumn = ToColumn,
                        //Set IsActive to false, and update later
                        IsActive = false
                    };

                    //Check if Relationship is active
                    foreach (AMO.MeasureGroupDimension MeasureGroupDimension in AMODatabase.Cubes[0].MeasureGroups.GetByName(FromTable.Name).Dimensions)
                        if ( MeasureGroupDimension is AMO.ReferenceMeasureGroupDimension)
                            if(((AMO.ReferenceMeasureGroupDimension)MeasureGroupDimension).RelationshipID == AMORelationship.ID)
                                TOMRelationship.IsActive = true;

                    //Add the Relationship to the Database
                    TOMDatabase.Model.Relationships.Add(TOMRelationship);
                    
                }
            }
            #endregion
            //Add TabularEditor compatibility annotation if necessary
            if (AddTabularEditorAnnotation)
                TOMDatabase.Model.Annotations.Add(new TOM.Annotation {
                    Name = "TabularEditor_CompatibilityVersion",
                    Value = AMODatabase.CompatibilityLevel.ToString()
                });
            //TODO: Handle KPIs
            return TOMDatabase;
        }
    }
}
