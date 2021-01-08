// DataExtractor is an ArcGIS add-in used to extract biodiversity
// information from SQL Server based on existing boundaries.
//
// Copyright © 2017 SxBRC, 2017-2019 TVERC, 2020 Andy Foy Consulting
//
// This file is part of DataExtractor.
//
// DataExtractor is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataExtractor is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with DataExtractor.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Windows.Forms;

using DataExtractor.Properties;
using HLFileFunctions;
using HLStringFunctions;


namespace HLExtractorToolConfig
{
    class ExtractorToolConfig
    {
        // Declare all the variables.
        // Environment and menu variables.
        bool DebugMode;
        string LogFilePath;
        string FileDSN;
        string ConnectionString;
        int TimeoutSeconds;
        string DefaultPath;
        string PartnerFolder;
        string DatabaseSchema;
        //string TableListSQL;
        string IncludeWildcard;
        string ExcludeWildcard;
        string PartnerTable;
        string PartnerColumn;
        string ShortColumn;
        string NotesColumn;
        string ActiveColumn;
        string FormatColumn;
        string ExportColumn;
        //string FilesColumn;
        string SQLFilesColumn;
        string SQLTableColumn;
        string MapFilesColumn;
        string TagsColumn;
        string SpatialColumn;
        string PartnerClause;
        List<string> SelectTypeOptions = new List<string>();
        int DefaultSelectType;
        // string RecMax; // Not sure we need this.
        bool DefaultZip;
        string ExclusionClause;
        bool DefaultExclusion;
        bool DefaultClearLogFile;
        bool DefaultUseCentroids;
        bool HideUseCentroids;
        bool DefaultUploadToServer;
        bool HideUploadToServer;
        
        // Layer variables - SQL.
        List<string> SQLTables = new List<string>();
        List<string> SQLOutputNames = new List<string>();
        List<string> SQLOutputTypes = new List<string>();
        List<string> SQLColumns = new List<string>();
        List<string> SQLWhereClauses = new List<string>();
        List<string> SQLOrderClauses = new List<string>();
        List<string> SQLMacroNames = new List<string>();
        List<string> SQLMacroParms = new List<string>();

        // Layer variables - Map layers.
        List<string> MapTables = new List<string>();
        List<string> MapTableNames = new List<string>();
        List<string> MapOutputNames = new List<string>();
        List<string> MapOutputTypes = new List<string>();
        List<string> MapColumns = new List<string>();
        List<string> MapWhereClauses = new List<string>();
        List<string> MapOrderClauses = new List<string>();
        List<string> MapMacroNames = new List<string>();
        List<string> MapMacroParms = new List<string>();

        bool FoundXML;
        bool LoadedXML;

        FileFunctions myFileFuncs;
        StringFunctions myStringFuncs;
        XmlElement xmlDataExtract;

        public ExtractorToolConfig(string anXMLProfile)
        {
            // Open XML
            myFileFuncs = new FileFunctions();
            myStringFuncs = new StringFunctions();
            string strXMLFile = anXMLProfile; // The user has specified this and we've checked it exists.
            FoundXML = true; // In this version we have already checked that it exists.
            LoadedXML = true;

            // Now get all the config variables.
            // Read the file.
            if (FoundXML)
            {
                XmlDocument xmlConfig = new XmlDocument();
                try
                {
                    xmlConfig.Load(strXMLFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error in XML file; cannot load. System error message: " + ex.Message, "XML Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }
                string strRawText;
                XmlNode currNode = xmlConfig.DocumentElement.FirstChild; // This gets us the DataExtractor.
                xmlDataExtract = (XmlElement)currNode;

                // XML loaded successfully; get all of the detail in the Config object.

                try
                {
                    DebugMode = false;
                    string strDebugMode = xmlDataExtract["Debug"].InnerText;
                    if (strDebugMode.ToLower() == "yes" || strDebugMode.ToLower() == "y")
                        DebugMode = true;
                }
                catch
                {
                    DebugMode = false;
                }

                try
                {
                    LogFilePath = xmlDataExtract["LogFilePath"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'LogFilePath' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    FileDSN = xmlDataExtract["FileDSN"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'FileDSN' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    ConnectionString = xmlDataExtract["ConnectionString"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'ConnectionString' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    string strTimeout = xmlDataExtract["TimeoutSeconds"].InnerText;
                    bool blSuccess;

                    if (strTimeout != "")
                    {

                        blSuccess = int.TryParse(strTimeout, out TimeoutSeconds);
                        if (!blSuccess)
                        {
                            MessageBox.Show("The value entered for TimeoutSeconds in the XML file is not an integer number");
                            LoadedXML = false;
                        }
                        if (TimeoutSeconds < 0)
                        {
                            MessageBox.Show("The value entered for TimeoutSeconds in the XML file is negative");
                            LoadedXML = false;
                        }
                    }
                    else
                    {
                        TimeoutSeconds = 0; // None given.
                    }

                }
                catch
                {
                    TimeoutSeconds = 0; // We don't really care if it's not in because there's a default anyway.
                    return;
                }

                try
                {
                    DefaultPath = xmlDataExtract["DefaultPath"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultPath' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    PartnerFolder = xmlDataExtract["PartnerFolder"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'PartnerFolder' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    DatabaseSchema = xmlDataExtract["DatabaseSchema"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DatabaseSchema' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    IncludeWildcard = xmlDataExtract["IncludeWildcard"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'IncludeWildcard' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    ExcludeWildcard = xmlDataExtract["ExcludeWildcard"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'ExcludeWildcard' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }


                try
                {
                    PartnerTable = xmlDataExtract["PartnerTable"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'PartnerTable' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    PartnerColumn = xmlDataExtract["PartnerColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'PartnerColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    ShortColumn = xmlDataExtract["ShortColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'ShortColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    NotesColumn = xmlDataExtract["NotesColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'NotesColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    ActiveColumn = xmlDataExtract["ActiveColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'ActiveColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    FormatColumn = xmlDataExtract["FormatColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'FormatColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    ExportColumn = xmlDataExtract["ExportColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'ExportColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    SQLTableColumn = xmlDataExtract["SQLTableColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SQLTableColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    SQLFilesColumn = xmlDataExtract["SQLFilesColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SQLFilesColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    MapFilesColumn = xmlDataExtract["MapFilesColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'MapFilesColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    TagsColumn = xmlDataExtract["TagsColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'TagsColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    SpatialColumn = xmlDataExtract["SpatialColumn"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SpatialColumn' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    PartnerClause = xmlDataExtract["PartnerClause"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'PartnerClause' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    strRawText = xmlDataExtract["SelectTypeOptions"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SelectTypeOptions' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                // Process the selection type options.
                try
                {
                    char[] chrSplitCharacter = { ';' };
                    string[] liOptionList = strRawText.Split(chrSplitCharacter);
                    foreach (string strEntry in liOptionList)
                    {
                        SelectTypeOptions.Add(strEntry);
                    }
                }
                catch
                {
                    MessageBox.Show("Error parsing 'SelectTypeOptions' string. Please check for correct string formatting and placement of delimiters", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    strRawText = xmlDataExtract["DefaultSelectType"].InnerText;
                    double i;
                    bool blResult = Double.TryParse(strRawText, out i);
                    if (blResult && (int)i <= SelectTypeOptions.Count)
                        DefaultSelectType = (int)i - 1; // Note we are returning this as a 0 based index value.
                    else if ((int)i > SelectTypeOptions.Count)
                    {
                        MessageBox.Show("The entry for 'DefaultSelectType' in the XML document is larger than the number of items in the SelectTypeOptions", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }
                    else if (strRawText == "")
                        DefaultSelectType = -1;
                    else
                    {
                        MessageBox.Show("The entry for 'DefaultSelectType' in the XML document is not an integer number", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultSelectType' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    DefaultZip = false;
                    string strDefaultZip = xmlDataExtract["DefaultZip"].InnerText;
                    if (strDefaultZip.ToLower() == "yes" || strDefaultZip.ToLower() == "y")
                        DefaultZip = true;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultZip' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    ExclusionClause = xmlDataExtract["ExclusionClause"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'ExclusionClause' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    DefaultExclusion = false;
                    string strDefaultExclusion = xmlDataExtract["DefaultExclusion"].InnerText;
                    if (strDefaultExclusion.ToLower() == "yes" || strDefaultExclusion.ToLower() == "y")
                        DefaultExclusion = true;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultExclusion' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    DefaultClearLogFile = false;
                    string strDefaultClearLogFile = xmlDataExtract["DefaultClearLogFile"].InnerText;
                    if (strDefaultClearLogFile.ToLower() == "yes" || strDefaultClearLogFile.ToLower() == "y")
                        DefaultClearLogFile = true;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultClearLogFile' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    DefaultUseCentroids = false;
                    HideUseCentroids = false;
                    string strDefaultUseCentroids = xmlDataExtract["DefaultUseCentroids"].InnerText;
                    if (strDefaultUseCentroids.ToLower() == "yes" || strDefaultUseCentroids.ToLower() == "y")
                        DefaultUseCentroids = true;
                    if (strDefaultUseCentroids == "")
                        HideUseCentroids = true;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultUseCentroids' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    DefaultUploadToServer = false;
                    HideUploadToServer = false;
                    string strDefaultUploadToServer = xmlDataExtract["DefaultUploadToServer"].InnerText;
                    if (strDefaultUploadToServer.ToLower() == "yes" || strDefaultUploadToServer.ToLower() == "y")
                        DefaultUploadToServer = true;
                    if (strDefaultUploadToServer == "")
                        HideUploadToServer = true;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultUploadToServer' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                // Layer options. 
                // SQL layers first.
                // Firstly, get all the entries.
                XmlElement SQLFileCollection = null;
                try
                {
                    SQLFileCollection = xmlDataExtract["SQLTables"];
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SQLTables' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                // Now cycle through them.
                if (SQLFileCollection != null)
                {
                    foreach (XmlNode aNode in SQLFileCollection)
                    {
                        // Only process if not a comment
                        if (aNode.NodeType != XmlNodeType.Comment)
                        {

                            string strName = aNode.Name; // The name of the SQL layer, as included in the Files in the partner table.
                            SQLTables.Add(strName);

                            try
                            {
                                SQLOutputNames.Add(aNode["OutputName"].InnerText); // The OUTPUT name 
                            }
                            catch
                            {
                                MessageBox.Show("Could not locate the item 'OutputName' for SQL layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                LoadedXML = false;
                                return;
                            }

                            try
                            {
                                SQLOutputTypes.Add(aNode["OutputType"].InnerText); // The OUTPUT type 
                            }
                            catch
                            {
                                SQLOutputTypes.Add("");
                            }

                            try
                            {
                                SQLColumns.Add(aNode["Columns"].InnerText);
                            }
                            catch
                            {
                                MessageBox.Show("Could not locate the item 'Columns' for SQL layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                LoadedXML = false;
                                return;
                            }

                            try
                            {
                                SQLWhereClauses.Add(aNode["WhereClause"].InnerText);
                            }
                            catch
                            {
                                SQLWhereClauses.Add("");
                            }

                            try
                            {
                                SQLOrderClauses.Add(aNode["OrderClause"].InnerText);
                            }
                            catch
                            {
                                SQLOrderClauses.Add("");
                            }

                            try
                            {
                                SQLMacroNames.Add(aNode["MacroName"].InnerText);
                            }
                            catch
                            {
                                SQLMacroNames.Add("");
                            }

                            try
                            {
                                SQLMacroParms.Add(aNode["MacroParm"].InnerText);
                            }
                            catch
                            {
                                SQLMacroParms.Add("");
                            }
                        }
                    }
                }

                // Now do the GIS Layers.
                XmlElement MapTableCollection = null;
                try
                {
                    MapTableCollection = xmlDataExtract["MapTables"];
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'MapTables' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                // Now cycle through them.
                if (MapTableCollection != null)
                {
                    foreach (XmlNode aNode in MapTableCollection)
                    {
                        // Only process if not a comment
                        if (aNode.NodeType != XmlNodeType.Comment)
                        {

                            string strName = aNode.Name; // The output name of the GIS layer (subset).
                            //strName = strName.Replace("_", " "); // Replace any underscores with spaces for better display.

                            MapTables.Add(strName);

                            try
                            {
                                MapTableNames.Add(aNode["TableName"].InnerText); // The name of the source layer.
                            }
                            catch
                            {
                                MessageBox.Show("Could not locate the item 'OutputName' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                LoadedXML = false;
                                return;
                            }

                            try
                            {
                                MapOutputNames.Add(aNode["OutputName"].InnerText); // The output name.
                            }
                            catch
                            {
                                MessageBox.Show("Could not locate the item 'OutputName' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                LoadedXML = false;
                                return;
                            }

                            try
                            {
                                MapOutputTypes.Add(aNode["OutputType"].InnerText); // The output type.
                            }
                            catch
                            {
                                MapOutputTypes.Add("");
                            }

                            try
                            {
                                MapColumns.Add(aNode["Columns"].InnerText); // Columns don't need to be formatted.
                            }
                            catch
                            {
                                MessageBox.Show("Could not locate the item 'Columns' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                LoadedXML = false;
                                return;
                            }

                            try
                            {
                                MapWhereClauses.Add(aNode["WhereClause"].InnerText);
                            }
                            catch
                            {
                                MapWhereClauses.Add("");
                            }

                            try
                            {
                                MapOrderClauses.Add(aNode["OrderClause"].InnerText);
                            }
                            catch
                            {
                                MapOrderClauses.Add("");
                            }

                            try
                            {
                                MapMacroNames.Add(aNode["Macro"].InnerText);
                            }
                            catch
                            {
                                MapMacroNames.Add("");
                            }

                            try
                            {
                                MapMacroParms.Add(aNode["MacroParm"].InnerText);
                            }
                            catch
                            {
                                MapMacroParms.Add("");
                            }
                        }
                    }
                }
            }
        }

        // Below here all the retrieval functions for the information.

        // 1. Success criteria.
        public bool GetFoundXML()
        {
            return FoundXML;
        }

        public bool GetLoadedXML()
        {
            return LoadedXML;
        }

        // General parameters.
        public bool GetDebugMode()
        {
            return DebugMode;
        }

        public string GetLogFilePath()
        {
            return LogFilePath;
        }

        public string GetSDEName()
        {
            return FileDSN;
        }

        public string GetConnectionString()
        {
            return ConnectionString;
        }

        public int GetTimeoutSeconds()
        {
            return TimeoutSeconds;
        }

        public string GetDefaultPath()
        {
            return DefaultPath;
        }

        public string GetPartnerFolder()
        {
            return PartnerFolder;
        }

        public string GetDatabaseSchema()
        {
            return DatabaseSchema;
        }

        //public string GetTableListSQL()
        //{
        //    return TableListSQL;
        //}
        public string GetIncludeWildcard()
        {
            return IncludeWildcard;
        }

        public string GetExcludeWildcard()
        {
            return ExcludeWildcard;
        }

        public string GetPartnerTable()
        {
            return PartnerTable;
        }
        public List<string> GetAllPartnerColumns()
        {
            List<string> AllPartnerColumns = new List<string>();
            AllPartnerColumns.Add(PartnerColumn);
            AllPartnerColumns.Add(ShortColumn);
            AllPartnerColumns.Add(NotesColumn);
            AllPartnerColumns.Add(ActiveColumn);
            AllPartnerColumns.Add(FormatColumn);
            AllPartnerColumns.Add(ExportColumn);
            AllPartnerColumns.Add(SQLTableColumn);
            AllPartnerColumns.Add(SQLFilesColumn);
            AllPartnerColumns.Add(MapFilesColumn);
            AllPartnerColumns.Add(TagsColumn);
            return AllPartnerColumns;
        }

        public string GetPartnerColumn()
        {
            return PartnerColumn;
        }

        public string GetShortColumn()
        {
            return ShortColumn;
        }

        public string GetNotesColumn()
        {
            return NotesColumn;
        }

        public string GetActiveColumn()
        {
            return ActiveColumn;
        }

        public string GetFormatColumn()
        {
            return FormatColumn;
        }

        public string GetExportColumn()
        {
            return ExportColumn;
        }

        public string GetSQLTableColumn()
        {
            return SQLTableColumn;
        }

        public string GetSQLFilesColumn()
        {
            return SQLFilesColumn;
        }

        public string GetMapFilesColumn()
        {
            return MapFilesColumn;
        }


        public string GetTagsColumn()
        {
            return TagsColumn;
        }

        public string GetSpatialColumn()
        {
            return SpatialColumn;
        }

        public string GetPartnerClause()
        {
            return PartnerClause;
        }

        public List<string> GetSelectTypeOptions()
        {
            return SelectTypeOptions;
        }

        public int GetDefaultSelectType()
        {
            return DefaultSelectType;
        }

        public bool GetDefaultZip()
        {
            return DefaultZip;
        }

        public string GetExclusionClause()
        {
            return ExclusionClause;
        }

        public bool GetDefaultExclusion()
        {
            return DefaultExclusion;
        }

        public bool GetDefaultClearLogFile()
        {
            return DefaultClearLogFile;
        }

        public bool GetDefaultUseCentroids()
        {
            return DefaultUseCentroids;
        }

        public bool GetHideUseCentroids()
        {
            return HideUseCentroids;
        }

        public bool GetDefaultUploadToServer()
        {
            return DefaultUploadToServer;
        }

        public bool GetHideUploadToServer()
        {
            return HideUploadToServer;
        }

        // 2. Layer variables - SQL.
        public List<string> GetSQLTables()
        {
            return SQLTables; // These are the subsets.
        }

        public List<string> GetSQLOutputNames()
        {
            return SQLOutputNames;
        }

        public List<string> GetSQLOutputTypes()
        {
            return SQLOutputTypes;
        }

        public List<string> GetUniqueSQLOutputNames()
        {
            List<string> liUniqueSQLNames = new List<string>();
            foreach (string strName in SQLOutputNames)
            {
                if (!liUniqueSQLNames.Contains(strName, StringComparer.OrdinalIgnoreCase))
                    liUniqueSQLNames.Add(strName);
            }
            return liUniqueSQLNames;
        }

        public List<string> GetSQLColumns()
        {
            return SQLColumns;
        }

        public List<string> GetSQLWhereClauses()
        {
            return SQLWhereClauses;
        }

        public List<string> GetSQLOrderClauses()
        {
            return SQLOrderClauses;
        }

        public List<string> GetSQLMacroNames()
        {
            return SQLMacroNames;
        }

        public List<string> GetSQLMacroParms()
        {
            return SQLMacroParms;
        }

        // 3. Layer variables - Map layers.

        public List<string> GetMapTables()
        {
            return MapTables; // The names of the subsets.
        }

        public List<string> GetMapTableNames()
        {
            return MapTableNames; // The names of the source layers.
        }

        public List<string> GetUniqueMapTableNames()
        {
            List<string> liUniqueMapNames = new List<string>();
            foreach (string strName in MapTableNames)
            {
                if (!liUniqueMapNames.Contains(strName, StringComparer.OrdinalIgnoreCase))
                    liUniqueMapNames.Add(strName);
            }
            return liUniqueMapNames;
        }

        public List<string> GetMapOutputNames()
        {
            return MapOutputNames; // The names of the outputs.
        }

        public List<string> GetMapOutputTypes()
        {
            return MapOutputTypes; // The type of the outputs.
        }

        public List<string> GetMapColumns()
        {
            return MapColumns;
        }

        public List<string> GetMapWhereClauses()
        {
            return MapWhereClauses;
        }

        public List<string> GetMapOrderClauses()
        {
            return MapOrderClauses;
        }

        public List<string> GetMapMacroNames()
        {
            return MapMacroNames;
        }

        public List<string> GetMapMacroParms()
        {
            return MapMacroParms;
        }


        private string GetConfigFilePath()
        {
            // Create folder dialog.
            FolderBrowserDialog xmlFolder = new FolderBrowserDialog();

            // Set the folder dialog title.
            xmlFolder.Description = "Select folder containing 'DataExtractor.xml' file ...";
            xmlFolder.ShowNewFolderButton = false;

            // Show folder dialog.
            if (xmlFolder.ShowDialog() == DialogResult.OK)
            {
                // Return the selected path.
                return xmlFolder.SelectedPath;
            }
            else
                return null;
        }
    }
}
