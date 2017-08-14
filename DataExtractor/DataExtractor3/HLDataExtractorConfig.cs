using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Windows.Forms;

using DataExtractor3.Properties;
using HLFileFunctions;
using HLStringFunctions;


namespace HLDataExtractorConfig
{
    class DataExtractorConfig
    {
        // Declare all the variables.
        // Environment and menu variables.
        string LogFilePath;
        string FileDSN;
        string ConnectionString;
        int TimeoutSeconds;
        string DefaultPath;
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
        string MapFilesColumn;
        string TagsColumn;
        List<string> SelectTypeOptions = new List<string>();
        int DefaultSelectType;
        // string RecMax; // Not sure we need this.
        bool DefaultZip;
        string ConfidentialClause;
        bool DefaultConfidential;
        bool DefaultClearLogFile;
        
        // Layer variables - SQL.
        List<string> SQLTables = new List<string>();
        List<string> SQLTableNames = new List<string>();
        List<string> SQLColumns = new List<string>();
        List<string> SQLClauses = new List<string>();
        List<string> SQLSymbology = new List<string>(); // This will work differently from MapInfo.

        // Layer variables - Map layers.
        List<string> MapLayers = new List<string>();
        List<string> MapLayerNames = new List<string>();
        List<string> MapColumns = new List<string>();
        List<string> MapClauses = new List<string>();

        bool FoundXML;
        bool LoadedXML;

        FileFunctions myFileFuncs;
        StringFunctions myStringFuncs;
        XmlElement xmlDataExtract;

        public DataExtractorConfig()
        {
            // Open XML
            myFileFuncs = new FileFunctions();
            myStringFuncs = new StringFunctions();
            string strXMLFile = null;
            FoundXML = false;
            LoadedXML = true;

            try
            {
                // Get the XML file
                strXMLFile = Settings.Default.XMLFile;
                if (String.IsNullOrEmpty(strXMLFile) || (!myFileFuncs.FileExists(strXMLFile))) // Can't find it or doesn't exist
                {
                    // Prompt the user for the location of the file.
                    string strFolder = GetConfigFilePath();
                    if (!String.IsNullOrEmpty(strFolder))
                        strXMLFile = strFolder + @"\DataExtractor.xml";
                }

                // check the xml file path exists
                if (myFileFuncs.FileExists(strXMLFile))
                {
                    Settings.Default.XMLFile = strXMLFile;
                    Settings.Default.Save();
                    FoundXML = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error " + ex.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            // Go forth and obtain all information.
            // Firstly, read the file.
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
                XmlNode currNode = xmlConfig.DocumentElement.FirstChild; // This gets us the DataSelector.
                xmlDataExtract = (XmlElement)currNode;

                // XML loaded successfully; get all of the detail in the Config object.
                
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
                    DatabaseSchema = xmlDataExtract["DatabaseSchema"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DatabaseSchema' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                //try
                //{
                //    TableListSQL = xmlDataExtract["TableListSQL"].InnerText;
                //}
                //catch
                //{
                //    MessageBox.Show("Could not locate the item 'TableListSQL' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    LoadedXML = false;
                //    return;
                //}
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
                    DefaultClearLogFile = false;
                    string strDefaultClearLogFile = xmlDataExtract["DefaultClearLogFile"].InnerText;
                    if (strDefaultClearLogFile == "Yes")
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
                    DefaultZip = false;
                    string strDefaultZip = xmlDataExtract["DefaultZip"].InnerText;
                    if (strDefaultZip == "Yes")
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
                    ConfidentialClause = xmlDataExtract["ConfidentialClause"].InnerText;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'ConfidentialClause' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                try
                {
                    DefaultConfidential = false;
                    string strDefaultConfidential = xmlDataExtract["DefaultConfidential"].InnerText;
                    if (strDefaultConfidential == "Yes")
                        DefaultConfidential = true;
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'DefaultConfidential' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }


                // Layer options. 
                // SQL layers first.
                // Firstly, get all the entries.
                XmlElement SQLLayerCollection = null;
                try
                {
                    SQLLayerCollection = xmlDataExtract["SQLTables"];
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'SQLTables' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }

                // Now cycle through them.
                foreach(XmlNode aNode in SQLLayerCollection)
                {
                    string strName = aNode.Name; // The name of the SQL layer, as included in the Files in the partner table.
                    SQLTables.Add(strName);
                    try
                    {
                        SQLTableNames.Add(aNode["TableName"].InnerText); // The OUTPUT name 
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'TableName' for SQL layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
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
                        SQLClauses.Add(aNode["Clauses"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'Clauses' for SQL layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }

                    try
                    {
                        SQLSymbology.Add(aNode["Symbology"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'Symbology' for SQL layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
                    }
                }

                // Now do the GIS Layers.
                XmlElement MapLayerCollection = null;
                try
                {
                    MapLayerCollection = xmlDataExtract["MapLayers"];
                }
                catch
                {
                    MessageBox.Show("Could not locate the item 'MapLayers' in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadedXML = false;
                    return;
                }
                // Now cycle through them.
                foreach (XmlNode aNode in MapLayerCollection)
                {
                    string strName = aNode.Name; // The output name of the GIS layer (subset).
                    //strName = strName.Replace("_", " "); // Replace any underscores with spaces for better display.
                    MapLayers.Add(strName);
                    try
                    {
                        MapLayerNames.Add(aNode["LayerName"].InnerText); // The name of the source layer.
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'TableName' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
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
                        MapClauses.Add(aNode["Clause"].InnerText);
                    }
                    catch
                    {
                        MessageBox.Show("Could not locate the item 'Clause' for map layer " + strName + " in the XML file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadedXML = false;
                        return;
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

        public string GetConfidentialClause()
        {
            return ConfidentialClause;
        }

        public bool GetDefaultConfidential()
        {
            return DefaultConfidential;
        }

        public bool GetDefaultClearLogFile()
        {
            return DefaultClearLogFile;
        }

       
        // 2. Layer variables - SQL.
        public List<string> GetSQLTables()
        {
            return SQLTables; // These are the subsets.
        }

        public List<string> GetSQLTableNames()
        {
            return SQLTableNames;
        }

        public List<string> GetUniqueSQLTableNames()
        {
            List<string> liUniqueSQLNames = new List<string>();
            foreach (string strName in SQLTableNames)
            {
                if (!liUniqueSQLNames.Contains(strName))
                    liUniqueSQLNames.Add(strName);
            }
            return liUniqueSQLNames;
        }

        public List<string> GetSQLColumns()
        {
            return SQLColumns;
        }

        public List<string> GetSQLClauses()
        {
            return SQLClauses;
        }

        public List<string> GetSQLSymbology()
        {
            return SQLSymbology;
        }


        // 3. Layer variables - Map layers.

        public List<string> GetMapLayers()
        {
            return MapLayers; // The names of the subsets.
        }

        public List<string> GetMapLayerNames()
        {
            return MapLayerNames; // The names of the source layers.
        }

        public List<string> GetUniqueMapLayerNames()
        {
            List<string> liUniqueMapNames = new List<string>();
            foreach (string strName in MapLayerNames)
            {
                if (!liUniqueMapNames.Contains(strName))
                    liUniqueMapNames.Add(strName);
            }
            return liUniqueMapNames;
        }

        public List<string> GetMapColumns()
        {
            return MapColumns;
        }

        public List<string> GetMapClauses()
        {
            return MapClauses;
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
