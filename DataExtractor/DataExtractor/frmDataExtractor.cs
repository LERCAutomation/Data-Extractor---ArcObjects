// DataExtractor is an ArcGIS add-in used to extract biodiversity
// information from SQL Server based on existing boundaries.
//
// Copyright © 2017 SxBRC, 2017-2018 TVERC
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
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

using HLESRISQLServerFunctions;
using HLArcMapModule;
using HLFileFunctions;
using HLExtractorToolLaunchConfig;
using HLExtractorToolConfig;
using HLStringFunctions;
using DataExtractor.Properties;

using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.GeoDatabaseUI;

using System.Data.SqlClient;

using Archive;

namespace DataExtractor
{
    public partial class frmDataExtractor : Form
    {
        // Initialise the form.
        ExtractorToolConfig myConfig;
        FileFunctions myFileFuncs;
        StringFunctions myStringFuncs;
        ArcMapFunctions myArcMapFuncs;
        ArcSDEFunctions myArcSDEFuncs;
        SQLServerFunctions mySQLServerFuncs;
        ExtractorToolLaunchConfig myLaunchConfig;

        // Because it's a pain to create enumerable objects in C#, we are cheating a little with multiple synched lists.
        List<string> liOpenLayers = new List<string>();
        List<string> liOpenFiles = new List<string>();
        List<string> liOpenColumns = new List<string>();
        List<string> liOpenClauses = new List<string>();
        List<string> liOpenLayerEntries = new List<string>();

        List<string> liSQLLayers = new List<string>();
        List<string> liSQLOutputNames = new List<string>();
        List<string> liSQLColumns = new List<string>();
        List<string> liSQLClauses = new List<string>();
        //List<string> liSQLLayerFiles = new List<string>();

        string strUserID;
        string strConfigFile = "";

        bool blOpenForm; // Tracks all the way through whether the form should open.
        //bool blFormHasOpened; // Informs all controls whether the form has opened.

        public frmDataExtractor()
        {

            blOpenForm = true; // Assume we're going to open
            //blFormHasOpened = false; // But we haven't yet.

            InitializeComponent();

            myLaunchConfig = new ExtractorToolLaunchConfig();
            myFileFuncs = new FileFunctions();

            if (!myLaunchConfig.XMLFound)
            {
                MessageBox.Show("XML file 'DataExtractor.xml' not found; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }
            if (!myLaunchConfig.XMLLoaded)
            {
                MessageBox.Show("Error loading XML File 'DataExtractor.xml'; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }

            if (blOpenForm)
            {
                string strXMLFolder = myFileFuncs.GetDirectoryName(Settings.Default.XMLFile);
                bool blOnlyDefault = true;
                int intCount = 0;
                if (myLaunchConfig.ChooseConfig) // If we are allowed to choose, check if there are multiple profiles. 
                // If there is only the default XML file in the directory, launch the form. Otherwise the user has to choose.
                {
                    foreach (string strFileName in myFileFuncs.GetAllFilesInDirectory(strXMLFolder))
                    {
                        if (myFileFuncs.GetFileName(strFileName).ToLower() != "dataextractor.xml" && myFileFuncs.GetExtension(strFileName).ToLower() == "xml")
                        {
                            // is it the default?
                            intCount++;
                            if (myFileFuncs.GetFileName(strFileName) != myLaunchConfig.DefaultXML)
                            {
                                blOnlyDefault = false;
                            }
                        }
                    }
                    if (intCount > 1)
                    {
                        blOnlyDefault = false;
                    }
                }
                if (myLaunchConfig.ChooseConfig && !blOnlyDefault)
                {
                    // User has to choose the configuration file first.

                    using (var myConfigForm = new frmChooseConfig(strXMLFolder, myLaunchConfig.DefaultXML))
                    {
                        var result = myConfigForm.ShowDialog();
                        if (result == System.Windows.Forms.DialogResult.OK)
                        {
                            strConfigFile = strXMLFolder + @"\" + myConfigForm.ChosenXMLFile;
                        }
                        else
                        {
                            MessageBox.Show("No XML file was chosen; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            blOpenForm = false;
                        }
                    }

                }
                else
                {
                    strConfigFile = strXMLFolder + @"\" + myLaunchConfig.DefaultXML; // don't allow the user to choose, just use the default.
                    // Just check it exists, though.
                    if (!myFileFuncs.FileExists(strConfigFile))
                    {
                        MessageBox.Show("The default XML file '" + myLaunchConfig.DefaultXML + "' was not found in the XML directory; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        blOpenForm = false;
                    }
                }
            }

            if (blOpenForm)
            {
                // Firstly let's read the XML.
                myConfig = new ExtractorToolConfig(strConfigFile); // Must now pass the correct XML name.

                IApplication pApp = ArcMap.Application;
                myArcMapFuncs = new ArcMapFunctions(pApp);
                myStringFuncs = new StringFunctions();

                // Did we find the XML?
                if (myConfig.GetFoundXML() == false)
                {
                    MessageBox.Show("XML file not found; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    blOpenForm = false;
                }

                // Did it load OK?
                else if (myConfig.GetLoadedXML() == false)
                {
                    MessageBox.Show("Error loading XML File; form cannot load.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    blOpenForm = false;
                }
            }

            // Close the form if there are any errors at this point.
            if (!blOpenForm)
            {
                Load += (s, e) => Close();
                return;
            }

            // Fix any illegal characters in the user name string
            strUserID = myStringFuncs.StripIllegals(Environment.UserName, "_", false);

            // The XML has loaded OK. Try to connect to the database and obtain the required info.
            myArcSDEFuncs = new ArcSDEFunctions();
            mySQLServerFuncs = new SQLServerFunctions();
            myFileFuncs = new FileFunctions();

            // Check if the output folders exist.
            if (!myFileFuncs.DirExists(myConfig.GetLogFilePath()))
            {
                MessageBox.Show("The log file path " + myConfig.GetLogFilePath() + " does not exist. Cannot load Data Extractor.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }

            if (!myFileFuncs.DirExists(myConfig.GetDefaultPath()) && blOpenForm)
            {
                MessageBox.Show("The output path " + myConfig.GetDefaultPath() + " does not exist. Cannot load Data Extractor.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }

            // Check if SDE file exists. If not, bail.
            string strSDE = myConfig.GetSDEName();
            if (!myFileFuncs.FileExists(strSDE) && blOpenForm)
            {
                MessageBox.Show("ArcSDE connection file " + strSDE + " not found. Cannot load Data Extractor.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                blOpenForm = false;
            }

            // SDE file exists; let's try to open it.
            IWorkspace wsSQLWorkspace = null;
            if (blOpenForm)
            {
                try
                {
                    wsSQLWorkspace = myArcSDEFuncs.OpenArcSDEConnection(strSDE);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot open ArcSDE connection " + strSDE + ". Error is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    blOpenForm = false;
                }
            }


            // Open the list of tables we can use.
            string strIncludeWildcard = myConfig.GetIncludeWildcard();
            string strExcludeWildcard = myConfig.GetExcludeWildcard();
            List<string> strTableList = myArcSDEFuncs.GetTableNames(wsSQLWorkspace, strIncludeWildcard, strExcludeWildcard);
            foreach (string strItem in strTableList)
            {
                lstTables.Items.Add(strItem);
            }

            // From the XML, get all the SQL tables and their items.
            liSQLLayers = myConfig.GetSQLTables(); //Node names as defined in the SQLFileList.
            liSQLOutputNames = myConfig.GetSQLTableNames(); // Output names for each node.
            liSQLColumns = myConfig.GetSQLColumns();
            liSQLClauses = myConfig.GetSQLClauses();
            //liSQLLayerFiles = myConfig.GetSQLSymbology();

            string strPartnerTable = myConfig.GetPartnerTable();
            string strPartnerColumn = myConfig.GetPartnerColumn();
            string strActiveColumn = myConfig.GetActiveColumn();

            if (blOpenForm)
            {
                // Open the list of partners.
                

                // Make sure the table and columns exist.
                if (!myArcMapFuncs.TableExists(strSDE + @"\" + myConfig.GetDatabaseSchema() + "." + strPartnerTable))
                {
                    MessageBox.Show("Partner table " + strPartnerTable + " not found. Cannot load Data Extractor.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    blOpenForm = false;
                }

                List<string> liAllPartnerColumns = myConfig.GetAllPartnerColumns();
                List<string> liExistingPartnerColumns = myArcMapFuncs.FieldsExist(strSDE, myConfig.GetDatabaseSchema() + "." + strPartnerTable, liAllPartnerColumns);
                var theDifference = liAllPartnerColumns.Except(liExistingPartnerColumns).ToList();
                if (theDifference.Count != 0)
                {
                    string strMessage = "The column(s) ";
                    foreach (string strCol in theDifference)
                    {
                        strMessage = strMessage + strCol + ", ";
                    }
                    strMessage = strMessage.Substring(0, strMessage.Length - 2) + "could not be found in table " + strPartnerTable + ". Cannot load Data Extractor."; 
                    MessageBox.Show(strMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    blOpenForm = false;
                }
            }

            if (blOpenForm)
            {
                IQueryFilter myFilter = new QueryFilter();
                myFilter.WhereClause = strActiveColumn + " = 'Y'";
                ITable tblPartnerTable = myArcMapFuncs.GetTable(wsSQLWorkspace, myConfig.GetDatabaseSchema() + "." + strPartnerTable);
                int intColumn = tblPartnerTable.FindField(strPartnerColumn);
                ICursor pCurs = tblPartnerTable.Search(myFilter, false);
                IRow pRow = null;
                List<string> strPartnerList = new List<string>();
                while ((pRow = pCurs.NextRow()) != null)
                {
                    lstActivePartners.Items.Add(pRow.get_Value(intColumn).ToString());
                }
                pCurs = null;
                myFilter = null;
            }

            // Now load all the map layers that are listed in the XML document into the menu, if they are loaded in the project.
            List<string> liAllLayers = myConfig.GetMapLayerNames(); // The GIS layer names.
            List<string> liAllFiles = myConfig.GetMapLayers(); // The node names; same as the Files in Partner table; used for output.
            List<string> liAllColumns = myConfig.GetMapColumns();
            List<string> liAllClauses = myConfig.GetMapClauses();

            if (blOpenForm)
            {
                List<string> liMissingLayers = new List<string>();
                int a = 0;
                foreach (string strLayer in liAllLayers) // For all possible source layers.
                {
                    if (myArcMapFuncs.LayerExists(strLayer))
                    {
                        string strItemEntry = liAllFiles[a] + " --> " + strLayer;
                        lstLayers.Items.Add(strItemEntry);
                        // Add all the info for this item to the lists.
                        liOpenLayerEntries.Add(strItemEntry);
                        liOpenLayers.Add(strLayer);
                        liOpenFiles.Add(liAllFiles[a]);
                        liOpenColumns.Add(liAllColumns[a]);
                        liOpenClauses.Add(liAllClauses[a]);
                    }
                    else
                        liMissingLayers.Add(strLayer);
                    a++;
                }

                if ((liMissingLayers.Count == liAllLayers.Count) && (liAllLayers.Count != 0)) // There are no layers loaded.
                {
                    MessageBox.Show("There are no GIS layers loaded in the view.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (liMissingLayers.Count == 1)
                    MessageBox.Show("The following GIS layer is not loaded: " + liMissingLayers[0] + ".", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else if (liMissingLayers.Count > 0)
                {
                    string strMessage = "The following GIS layers are not loaded: ";
                    foreach (string strLayer in liMissingLayers)
                        strMessage = strMessage + strLayer + ", ";
                    strMessage = strMessage.Substring(0, strMessage.Length - 2) + ".";
                    MessageBox.Show(strMessage, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            // Finally set the check boxes and dropdown lists.
            if (blOpenForm)
            {
                chkZip.Checked = myConfig.GetDefaultZip();
                chkConfidential.Checked = myConfig.GetDefaultConfidential();
                chkClearLog.Checked = myConfig.GetDefaultClearLogFile();
                chkUseCentroids.Checked = myConfig.GetDefaultUseCentroids();

                foreach (string anOption in myConfig.GetSelectTypeOptions())
                    cmbSelectionType.Items.Add(anOption);

                if (myConfig.GetDefaultSelectType() != -1)
                    cmbSelectionType.SelectedIndex = myConfig.GetDefaultSelectType();
            }

            // Hide controls that were not requested
            if (myConfig.GetConfidentialClause() == "")
            {
                chkConfidential.Hide();
            }
            else
            {
                chkConfidential.Show();
            }

            if (myConfig.GetHideUseCentroids())
            {
                chkUseCentroids.Hide();
            }
            else
            {
                chkUseCentroids.Show();
            }

            // tidy up.
            wsSQLWorkspace = null;
        
            // Finally decide whether the form should open at all and return control to the application.
            if (!blOpenForm)
            {
                Load += (s, e) => Close();
                return;
            }
            
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            // User has clicked OK. 
            this.Cursor = Cursors.WaitCursor;

            // Firstly check all the information. If anything doesn't add up, bail out.

            // Have they selected at least one partner?
            if (lstActivePartners.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one partner.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            // Have they selected either a table or a GIS layer?
            if (lstLayers.SelectedItems.Count == 0 && lstTables.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select one or more layers.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            // Have they selected an output type?
            if (cmbSelectionType.SelectedItem.ToString() == "")
            {
                MessageBox.Show("Please choose a selection type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.BringToFront();
                this.Cursor = Cursors.Default;
                return;
            }

            // So far so good. Start the process.

            //================================================ PROCESS ===========================================================
            
            // Get the selected partners.
            List<string> liChosenPartners = new List<string>();
            foreach (string aPartner in lstActivePartners.SelectedItems)
            {
                liChosenPartners.Add(aPartner);
            }


            string strChosenSQLLayer = "";
            if (lstTables.SelectedItem != null)
                strChosenSQLLayer = lstTables.SelectedItem.ToString();

            // Get the selected GIS layers
            List<string> liChosenGISLayers = new List<string>();
            List<string> liChosenGISFiles = new List<string>();
            List<string> liChosenGISColumns = new List<string>();
            List<string> liChosenGISClauses = new List<string>();
            List<string> liChosenGISLayerEntries = new List<string>();

            foreach (string aGISLayer in lstLayers.SelectedItems)
            {
                liChosenGISLayerEntries.Add(aGISLayer);
                // Find this entry to get the index.
                int a = liOpenLayerEntries.IndexOf(aGISLayer);
                // It always exists because we've loaded it ourselves. Populate the other lists.
                liChosenGISLayers.Add(liOpenLayers[a]);
                liChosenGISFiles.Add(liOpenFiles[a]);
                liChosenGISColumns.Add(liOpenColumns[a]);
                liChosenGISClauses.Add(liOpenClauses[a]);
            }

            string strSelectionType = cmbSelectionType.SelectedItem.ToString();
            bool blCreateZip = chkZip.Checked;
            bool blConfidential = chkConfidential.Checked;
            bool blClearLogFile = chkClearLog.Checked;
            bool blUseCentroids = chkUseCentroids.Checked;

            string strUseCentroids = "0"; // Use polygons is the default;
            if (blUseCentroids)
                strUseCentroids = "1";

            string strSelectionDigit = "1"; // Spatial is the default.
            // Check the selection type and set the appropriate number.
            if (strSelectionType.ToLower().Contains("spatial"))
            {
                if (strSelectionType.ToLower().Contains("tags"))
                {
                    strSelectionDigit = "3";
                }
            }
            else
            {
                strSelectionDigit = "2";
            }

            // Everything has been taken from the menu. Get some further data from the Config file.
            string strLogFilePath = myConfig.GetLogFilePath();
            string strDefaultPath = myConfig.GetDefaultPath();
            string strSDEName = myConfig.GetSDEName();
            string strConnectionString = myConfig.GetConnectionString();
            string strDatabaseSchema = myConfig.GetDatabaseSchema();

            string strPartnerTable = myConfig.GetPartnerTable();
            string strPartnerColumn = myConfig.GetPartnerColumn();
            string strShortColumn = myConfig.GetShortColumn();
            string strNotesColumn = myConfig.GetNotesColumn();
            string strFormatColumn = myConfig.GetFormatColumn();
            string strExportColumn = myConfig.GetExportColumn();
            string strSQLFilesColumn = myConfig.GetSQLFilesColumn();
            string strMapFilesColumn = myConfig.GetMapFilesColumn();
            string strTagsColumn = myConfig.GetTagsColumn();
            string strPartnerSpatialColumn = myConfig.GetSpatialColumn();

            string strConfidentialClause = myConfig.GetConfidentialClause();

            // Finally get all the raw data for the SQL layers and maps into memory.
            // Note this has changed - we are processing *all* the SQL layers.

            List<string> liChosenSQLLayers = liSQLLayers; //new List<string>(); 
            List<string> liChosenSQLOutputNames = liSQLOutputNames; //new List<string>();
            List<string> liChosenSQLColumns = liSQLColumns; //new List<string>();
            List<string> liChosenSQLClauses = liSQLClauses; //new List<string>();
            //List<string> liChosenSQLSymbology = liSQLLayerFiles; //new List<string>();

            //int b = 0;
            //foreach (string strTable in liSQLOutputNames)
            //{
            //    if (strTable == strChosenSQLLayer)
            //    {
            //        liChosenSQLLayers.Add(liSQLLayers[b]);
            //        liChosenSQLOutputNames.Add(strTable);
            //        liChosenSQLColumns.Add(liSQLColumns[b]);
            //        liChosenSQLClauses.Add(liSQLClauses[b]);
            //        liChosenSQLSymbology.Add(liSQLLayerFiles[b]);
            //    }
            //    b++;
            //}

            // Start the log file.
            string strLogFile = strLogFilePath + @"\DataExtractor_" + strUserID + ".log";
            if (blClearLogFile)
            {
                bool blDeleted = myFileFuncs.DeleteFile(strLogFile);
                if (!blDeleted)
                {
                    MessageBox.Show("Cannot delete log file. Please make sure it is not open in another window");
                    this.BringToFront();
                    this.Cursor = Cursors.Default;
                    return;
                }
            }

            myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");
            myFileFuncs.WriteLine(strLogFile, "Process started");
            myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");

            myArcMapFuncs.ToggleDrawing();
            myArcMapFuncs.ToggleTOC();
            this.BringToFront();

            if (strUserID == "Temp")
                myFileFuncs.WriteLine(strLogFile, "Please note user ID is: 'Temp'");
            else
                myFileFuncs.WriteLine(strLogFile, "User ID is: '" + strUserID + "'");

            // Load the partner table if it's not already there. We know it exists.
            if (!myArcMapFuncs.LayerExists(strPartnerTable))
            {
                IFeatureClass pFC = myArcMapFuncs.GetFeatureClass(strSDEName, myConfig.GetDatabaseSchema() + "." + strPartnerTable);
                bool blResult = myArcMapFuncs.AddLayerFromFClass(pFC);
                pFC = null;
                if (!blResult)
                {
                    MessageBox.Show("Cannot add partner table to ArcGIS. Aborting.");
                    myFileFuncs.WriteLine(strLogFile, "Cannot add partner table to ArcGIS. Aborting.");
                    this.BringToFront();
                    this.Cursor = Cursors.Default;
                    return;
                }
            }

            myFileFuncs.WriteLine(strLogFile, "Species table is: '" + strChosenSQLLayer + "'");
            if (strUseCentroids == "1")
                myFileFuncs.WriteLine(strLogFile, "Polygons will be selected using centroids.");
            else
                myFileFuncs.WriteLine(strLogFile, "Polygons will be selected using boundary.");
                                 
            // Now process the selected partners.
            foreach (string strPartner in liChosenPartners)
            {
                lblPartner.Text = "Partner: " + strPartner;
                lblPartner.Visible = true;
                lblPartner.Refresh();

                // Variables that will be filled.
                string strShortName = null;
                string strNotes = null;
                string strFormat = null;
                string strExport = null;
                List<string> liSQLFilesRaw = null;
                List<string> liMapFilesRaw = null;
                List<string> liSQLFiles = new List<string>();
                List<string> liMapFiles = new List<string>();
                string strTags = null;

                myFileFuncs.WriteLine(strLogFile, "");
                myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");
                myFileFuncs.WriteLine(strLogFile, "Processing partner " + strPartner + ".");

                // Get all the information on this partner. Firstly select the correct polygon.
                myFileFuncs.WriteLine(strLogFile, "Selecting partner boundary and retrieving information for partner " + strPartner + ".");
                string strWhereClause = strPartnerColumn + " = '" + strPartner + "'";
                myArcMapFuncs.SelectLayerByAttributes(strPartnerTable, strWhereClause);

                // Get the associated cursor.
                ICursor pCurs = myArcMapFuncs.GetCursorOnFeatureLayer(strPartnerTable);

                // Extract the information and report any strangeness.
                if (pCurs == null)
                {
                    MessageBox.Show("Could not retrieve information for partner " + strPartner + ". Aborting.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    myFileFuncs.WriteLine(strLogFile, "Could not retrieve information for partner " + strPartner + ". Aborting.");
                    this.BringToFront();
                    this.Cursor = Cursors.Default;
                    return;
                }

                IRow pRow = null;
                int i = 0;
                while ((pRow = pCurs.NextRow()) != null)
                {
                    int a;
                    if (i == 0) // Only take the first entry.
                    {
                        a = pRow.Fields.FindField(strShortColumn);
                        strShortName = pRow.get_Value(a).ToString();
                        a = pRow.Fields.FindField(strNotesColumn);
                        strNotes = pRow.get_Value(a).ToString();
                        a = pRow.Fields.FindField(strFormatColumn);
                        strFormat = pRow.get_Value(a).ToString();
                        a = pRow.Fields.FindField(strExportColumn);
                        strExport = pRow.get_Value(a).ToString();
                        a = pRow.Fields.FindField(strSQLFilesColumn);
                        liSQLFilesRaw = pRow.get_Value(a).ToString().Split(',').ToList(); // List of the different SQL file names.
                        a = pRow.Fields.FindField(strMapFilesColumn);
                        liMapFilesRaw = pRow.get_Value(a).ToString().Split(',').ToList(); // List of the different Map file names.
                        a = pRow.Fields.FindField(strTagsColumn);
                        strTags = pRow.get_Value(a).ToString();

                        // sort out any spaces in the files list.
                        foreach (string aFile in liSQLFilesRaw)
                        {
                            liSQLFiles.Add(aFile.Trim());
                        }
                        foreach (string aFile in liMapFilesRaw)
                        {
                            liMapFiles.Add(aFile.Trim());
                        }
                    }
                    i++;
                }

                if (i > 1)
                {
                    MessageBox.Show("There are duplicate entries for partner " + strPartner + " in the partner table. Using the first entry.");
                    myFileFuncs.WriteLine(strLogFile, "There are duplicate entries for partner " + strPartner + " in the partner table. Using the first entry.");
                }

                // Create the connection.
                SqlConnection dbConn = mySQLServerFuncs.CreateSQLConnection(myConfig.GetConnectionString());

                // Note under current setup the CurrentSpatialTable never changes.
                if (strChosenSQLLayer != "")
                {
                    myFileFuncs.WriteLine(strLogFile, "Processing SQL layers for partner " + strPartner + ".");
                    int b = 0;
                    int intTotLayers = liChosenSQLLayers.Count(); // This is *all* the SQL layers in the XML.
                    //int intThisLayer = 1;

                    // Firstly do the spatial/ tags selection.
                    int intTimeoutSeconds = myConfig.GetTimeoutSeconds();
                    string strIntermediateTable = strChosenSQLLayer + "_" + strUserID; // The output table from the HLSppSelection stored procedure.

                    // Delete the original subset (in case it still exists).
                    SqlCommand myDeleteCommand = mySQLServerFuncs.CreateSQLCommand(ref dbConn, "HLClearSpatialSubset", CommandType.StoredProcedure);
                    mySQLServerFuncs.AddSQLParameter(ref myDeleteCommand, "Schema", strDatabaseSchema);
                    mySQLServerFuncs.AddSQLParameter(ref myDeleteCommand, "SpeciesTable", strChosenSQLLayer);
                    mySQLServerFuncs.AddSQLParameter(ref myDeleteCommand, "UserId", strUserID);
                    try
                    {
                        //myFileFuncs.WriteLine(strLogFile, "Opening SQL Connection");
                        dbConn.Open();
                        myFileFuncs.WriteLine(strLogFile, "Executing stored procedure to delete spatial subselection.");
                        string strRowsAffect = myDeleteCommand.ExecuteNonQuery().ToString();
                        // Close the connection again.
                        myDeleteCommand = null;
                        dbConn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Could not execute stored procedure 'HLClearSpatialSubset'. System returned the following message: " +
                            ex.Message);
                        myFileFuncs.WriteLine(strLogFile, "Could not execute stored procedure HLClearSpatialSubset. System returned the following message: " +
                            ex.Message);
                        this.Cursor = Cursors.Default;
                        dbConn.Close();
                        myArcMapFuncs.ToggleDrawing();
                        myArcMapFuncs.ToggleTOC();
                        lblPartner.Text = "";
                        this.BringToFront();
                        this.Cursor = Cursors.Default;
                        return;
                    }

                    // make the spatial / tags selection.
                    string strStoredSpatialAndTagsProcedure = "AFSelectSppRecords";
                    SqlCommand mySpatialCommand = null;
                    if (intTimeoutSeconds == 0)
                    {
                        mySpatialCommand = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strStoredSpatialAndTagsProcedure, CommandType.StoredProcedure); // Note pass connection by ref here.
                    }
                    else
                    {
                        mySpatialCommand = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strStoredSpatialAndTagsProcedure, CommandType.StoredProcedure, intTimeoutSeconds);
                    }

                    mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "Schema", strDatabaseSchema);
                    mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "PartnerTable", strPartnerTable);
                    mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "PartnerColumn", strShortColumn); // Used for selection
                    mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "Partner", strShortName); // Used for selection
                    mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "TagsColumn", strTagsColumn);
                    mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "PartnerSpatialColumn", strPartnerSpatialColumn);
                    mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "SelectType", strSelectionDigit);
                    mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "SpeciesTable", strChosenSQLLayer);
                    mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "UserId", strUserID);
                    mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "UseCentroids", strUseCentroids);
                    //mySQLServerFuncs.AddSQLParameter(ref mySpatialCommand, "Split", strSplit);

                    // Execute stored procedure.
                    int intCount = 0;
                    bool blSuccess = true;
                    try
                    {
                        //myFileFuncs.WriteLine(strLogFile, "Opening SQL Connection");
                        dbConn.Open();
                        myFileFuncs.WriteLine(strLogFile, "Executing stored procedure to make spatial / tags selection.");
                        string strRowsAffect = mySpatialCommand.ExecuteNonQuery().ToString();

                        blSuccess = mySQLServerFuncs.TableHasRows(ref dbConn, strIntermediateTable);
                        if (blSuccess)
                        {
                            intCount = mySQLServerFuncs.CountRows(ref dbConn, strIntermediateTable);


                            myFileFuncs.WriteLine(strLogFile, "Procedure returned " + intCount.ToString() + " records.");
                        }
                        else
                        {
                            myFileFuncs.WriteLine(strLogFile, "Procedure returned no records.");
                        }

                        // Close the connection again.
                        mySpatialCommand = null;
                        dbConn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Could not execute stored procedure. System returned the following message: " +
                            ex.Message);
                        myFileFuncs.WriteLine(strLogFile, "Could not execute stored procedure. System returned the following message: " +
                            ex.Message);
                        this.Cursor = Cursors.Default;
                        dbConn.Close();
                        myArcMapFuncs.ToggleDrawing();
                        myArcMapFuncs.ToggleTOC();
                        lblPartner.Text = "";
                        this.BringToFront();
                        this.Cursor = Cursors.Default;
                        return;
                    }

                    foreach (string strSQLLayer in liChosenSQLLayers) // Output files, not input tables.
                    {
                        // Does the partner want this layer?
                        if (liSQLFiles.Contains(strSQLLayer))
                        {
                            // They do, process. 
                            lblPartner.Text = strPartner + ": Processing SQL table " + strSQLLayer + ".";
                            lblPartner.Refresh();

                            // Set up the final output.
                            string strOutFolder = strDefaultPath + @"\" + strShortName;

                            if (!myFileFuncs.DirExists(strOutFolder))
                            {
                                myFileFuncs.WriteLine(strLogFile, "Output folder " + strOutFolder + " doesn't exist. Creating...");
                                try
                                {
                                    Directory.CreateDirectory(strOutFolder);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Cannot create directory " + strOutFolder + ". System error: " + ex.Message);
                                    myArcMapFuncs.ToggleDrawing();
                                    myArcMapFuncs.ToggleTOC();
                                    lblPartner.Text = "";
                                    this.BringToFront();
                                    this.Cursor = Cursors.Default;
                                    return;
                                }
                                myFileFuncs.WriteLine(strLogFile, "Output folder created.");
                            }

                            bool blSpatial = false;
                            string strColumns = liChosenSQLColumns[b];
                            string strClause = liChosenSQLClauses[b];
                            string strOutputTable = liChosenSQLOutputNames[b];
                            //string strLayerFile = liChosenSQLSymbology[b];
                            myFileFuncs.WriteLine(strLogFile, "");
                            myFileFuncs.WriteLine(strLogFile, "Processing SQL layer " + strSQLLayer + " for output " + strOutputTable);

                            string strSpatialColumn = "";
                            string strSplit = "0"; // Do we need to split for polys / points?
                            // Is there a geometry field in the data requested?
                            string[] strGeometryFields = { "SP_GEOMETRY", "Shape" }; // Expand as required.
                            foreach (string strField in strGeometryFields)
                            {
                                if (strColumns.ToLower().Contains(strField.ToLower()))
                                {
                                    blSpatial = true;
                                    strSplit = "1"; // To be passed to stored procedure.
                                    strSpatialColumn = strField; // To be passed to stored procedure
                                }
                            }

                            // if '*' is used then check for geometry field in the table.

                            if (strColumns == "*")
                            {
                                string strCheckTable = strDatabaseSchema + "." + strChosenSQLLayer;
                                dbConn.Open();
                                foreach (string strField in strGeometryFields)
                                {
                                    if (mySQLServerFuncs.FieldExists(ref dbConn, strCheckTable, strField))
                                    {
                                        blSpatial = true;
                                        strSpatialColumn = strField;
                                        strSplit = "1";
                                    }
                                }
                                dbConn.Close();
                            }

                            // Do the attribute query. This splits the output into points and polygons as relevant.
                            // Set the temporary table names and the stored procedure names.
                            string strStoredProcedure = "HLSelectSppSubset"; // Default for all data
                            //string strPolyFC = strDatabaseSchema + "." + strIntermediateTable + "_poly"; // Change these.
                            //string strPointFC = strDatabaseSchema + "." + strIntermediateTable + "_point";
                            //string strTempTable = strDatabaseSchema + "." + strIntermediateTable + "_flat";
                            string strPolyFC = strIntermediateTable + "_poly"; // Change these.
                            string strPointFC = strIntermediateTable + "_point";
                            string strTempTable = strIntermediateTable + "_flat";

                            SqlCommand myCommand = null;

                            if (intTimeoutSeconds == 0)
                            {
                                myCommand = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strStoredProcedure, CommandType.StoredProcedure); // Note pass connection by ref here.
                            }
                            else
                            {
                                myCommand = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strStoredProcedure, CommandType.StoredProcedure, intTimeoutSeconds);
                            }

                            // Add the confidential clause if required.
                            if (strClause == "")
                            {
                                if (!blConfidential && strConfidentialClause != "") // Confidential records are excluded
                                {
                                    strClause = strConfidentialClause; // Note WHERE is already included in the SP and needs not be repeated.
                                }
                            }
                            else
                            {
                                if (!blConfidential && strConfidentialClause != "") // There is a where clause and confidential records are excluded
                                {
                                    strClause = strClause + " AND (" + strConfidentialClause + ")";
                                }
                            }

                            mySQLServerFuncs.AddSQLParameter(ref myCommand, "Schema", strDatabaseSchema);
                            mySQLServerFuncs.AddSQLParameter(ref myCommand, "SpeciesTable", strIntermediateTable);
                            mySQLServerFuncs.AddSQLParameter(ref myCommand, "SpatialColumn", strSpatialColumn);
                            mySQLServerFuncs.AddSQLParameter(ref myCommand, "ColumnNames", strColumns);
                            mySQLServerFuncs.AddSQLParameter(ref myCommand, "WhereClause", strClause);
                            mySQLServerFuncs.AddSQLParameter(ref myCommand, "GroupByClause", "");
                            mySQLServerFuncs.AddSQLParameter(ref myCommand, "OrderByClause", "");
                            mySQLServerFuncs.AddSQLParameter(ref myCommand, "UserID", strUserID);
                            mySQLServerFuncs.AddSQLParameter(ref myCommand, "Split", strSplit);

                            myFileFuncs.WriteLine(strLogFile, "Column names are: " + strColumns);
                            myFileFuncs.WriteLine(strLogFile, "Spatial column is: '" + strSpatialColumn + "'");
                            myFileFuncs.WriteLine(strLogFile, "Output base name is: '" + strIntermediateTable + "'");
                            if (strSplit == "1")
                                myFileFuncs.WriteLine(strLogFile, "Data is spatial and will be split into a point and a polygon layer.");
                            else
                                myFileFuncs.WriteLine(strLogFile, "Data is not spatial and will not be split.");
                            if (strClause.Length > 0)
                                myFileFuncs.WriteLine(strLogFile, "Where clause is: " + strClause.Replace("\r\n", " "));
                            else
                                myFileFuncs.WriteLine(strLogFile, "No where clause was used.");

                            // Open SQL connection to database and
                            // Run the stored procedure.
                            int intPolyCount = 0;
                            int intPointCount = 0;
                            try
                            {
                                //myFileFuncs.WriteLine(strLogFile, "Opening SQL Connection");
                                dbConn.Open();
                                myFileFuncs.WriteLine(strLogFile, "Executing stored procedure to make subset selection.");
                                string strRowsAffect = myCommand.ExecuteNonQuery().ToString();
                                if (blSpatial)
                                {
                                    blSuccess = mySQLServerFuncs.TableHasRows(ref dbConn, strPointFC);
                                    if (!blSuccess)
                                        blSuccess = mySQLServerFuncs.TableHasRows(ref dbConn, strPolyFC);
                                }
                                else
                                    blSuccess = mySQLServerFuncs.TableHasRows(ref dbConn, strTempTable);

                                if (blSuccess && blSpatial)
                                {
                                    intPolyCount = mySQLServerFuncs.CountRows(ref dbConn, strPolyFC);
                                    intPointCount = mySQLServerFuncs.CountRows(ref dbConn, strPointFC);


                                    myFileFuncs.WriteLine(strLogFile, "Procedure returned " + intPointCount.ToString() + " point and " + intPolyCount.ToString() +
                                        " polygon records.");
                                }
                                else if (blSuccess)
                                {
                                    intCount = mySQLServerFuncs.CountRows(ref dbConn, strTempTable);
                                    myFileFuncs.WriteLine(strLogFile, "Procedure returned " + intCount.ToString() + " records.");
                                }
                                else
                                {
                                    myFileFuncs.WriteLine(strLogFile, "Procedure returned no records.");
                                }


                                //myFileFuncs.WriteLine(strLogFile, "Closing SQL Connection");
                                dbConn.Close();
                                myCommand = null;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Could not execute stored procedure. System returned the following message: " +
                                    ex.Message);
                                myFileFuncs.WriteLine(strLogFile, "Could not execute stored procedure. System returned the following message: " +
                                    ex.Message);
                                this.Cursor = Cursors.Default;
                                dbConn.Close();
                                myArcMapFuncs.ToggleDrawing();
                                myArcMapFuncs.ToggleTOC();
                                this.BringToFront();
                                this.Cursor = Cursors.Default;
                                return;
                            }

                            lblPartner.Text = strPartner + ": Writing output for " + strSQLLayer + " to GIS format.";
                            lblPartner.Refresh();
                            this.BringToFront();

                            bool blIsGDB = false;
                            if (strFormat == "GDB")
                            {
                                blIsGDB = true;
                                strOutFolder = strOutFolder + "\\DataExtracts.gdb";
                                if (!myFileFuncs.DirExists(strOutFolder))
                                {
                                    myFileFuncs.WriteLine(strLogFile, "Output geodatabase " + strOutFolder + " doesn't exist. Creating...");
                                    myArcMapFuncs.CreateGeodatabase(strOutFolder);
                                }
                                myFileFuncs.WriteLine(strLogFile, "Output geodatabase created.");
                            }

                            string strOutLayer = strOutFolder + @"\" + strOutputTable; //strSQLLayer; //NOTE not poly or point - for non-spatial output.
                            string strOutLayerPoint = strOutLayer + "_point";
                            string strOutLayerPoly = strOutLayer + "_poly";


                            // Now export to shape or table as appropriate.
                            bool blResult = false;
                            if (blSpatial && blSuccess)
                            {
                                if (!blIsGDB) strOutLayerPoint = strOutLayerPoint + ".shp";
                                if (!blIsGDB) strOutLayerPoly = strOutLayerPoly + ".shp";
                                if (intPolyCount > 0)
                                {
                                    myFileFuncs.WriteLine(strLogFile, "Exporting polygon selection to GIS file: " + strOutLayerPoly + ".");
                                    //strPolyFC = strSDEName + @"\" + strPolyFC;
                                    blResult = myArcMapFuncs.CopyFeatures(strSDEName + @"\" + strDatabaseSchema + "." + strPolyFC, strOutLayerPoly); // Copies both to GDB and shapefile.
                                    myArcMapFuncs.RemoveLayer(strOutputTable + "_poly"); // Temporary layer is removed.
                                    if (!blResult)
                                    {
                                        MessageBox.Show("Error exporting polygon output from SQL table");
                                        myFileFuncs.WriteLine(strLogFile, "Error exporting " + strPolyFC + " to " + strOutLayerPoly);
                                        this.Cursor = Cursors.Default;
                                        myArcMapFuncs.ToggleDrawing();
                                        myArcMapFuncs.ToggleTOC();
                                        lblPartner.Text = "";
                                        this.BringToFront();
                                        this.Cursor = Cursors.Default;
                                        return;
                                    }
                                }
                                if (intPointCount > 0)
                                {
                                    myFileFuncs.WriteLine(strLogFile, "Exporting point selection to GIS file: " + strOutLayerPoint + ".");
                                    //strPointFC = strSDEName + @"\" + strPointFC;
                                    blResult = myArcMapFuncs.CopyFeatures(strSDEName + @"\" + strDatabaseSchema + "." + strPointFC, strOutLayerPoint); // Copies both to GDB and shapefile.
                                    myArcMapFuncs.RemoveLayer(strOutputTable + "_point"); // Temporary layer is removed.
                                    if (!blResult)
                                    {
                                        MessageBox.Show("Error exporting point output from SQL table");
                                        myFileFuncs.WriteLine(strLogFile, "Error exporting " + strPointFC + " to " + strOutLayerPoint);
                                        this.Cursor = Cursors.Default;
                                        myArcMapFuncs.ToggleDrawing();
                                        myArcMapFuncs.ToggleTOC();
                                        lblPartner.Text = "";
                                        this.BringToFront();
                                        this.Cursor = Cursors.Default;
                                        return;
                                    }
                                }

                            }
                            else if (blSuccess)
                            {
                                // There is only one table to export.
                                if (!blIsGDB) strOutLayer = strOutLayer + ".dbf";
                                //string strInTable = strSDEName + @"\" + strTempTable;
                                myFileFuncs.WriteLine(strLogFile, "Exporting selection to flat table: " + strOutLayer + ".");
                                blResult = myArcMapFuncs.CopyTable(strSDEName + @"\" + strDatabaseSchema + "." + strTempTable, strOutLayer);
                                if (!blResult)
                                {
                                    MessageBox.Show("Error exporting output from SQL table");
                                    myFileFuncs.WriteLine(strLogFile, "Error exporting " + strTempTable + " to " + strOutLayer);
                                    this.Cursor = Cursors.Default;
                                    myArcMapFuncs.ToggleDrawing();
                                    myArcMapFuncs.ToggleTOC();
                                    lblPartner.Text = "";
                                    this.BringToFront();
                                    this.Cursor = Cursors.Default;
                                    return;
                                }
                                myArcMapFuncs.RemoveStandaloneTable(strOutputTable);
                            }

                            lblPartner.Text = strPartner + ": Writing output for " + strSQLLayer + " to text file.";
                            lblPartner.Refresh();

                            // Now export to CSV if required.
                            strExport = strExport.ToLower().Trim();
                            if (strExport == "csv" && blSuccess)
                            {
                                if (blSpatial)
                                {

                                    bool blAppend = false;
                                    string strOutputFile = strDefaultPath + @"\" + strShortName + @"\" + strSQLLayer + ".csv";
                                    if (intPointCount > 0)
                                    {
                                        myFileFuncs.WriteLine(strLogFile, "Copying point results to CSV file " + strOutputFile);
                                        blResult = myArcMapFuncs.CopyToCSV(strOutLayerPoint, strOutputFile, true, blAppend, true);
                                        if (!blResult)
                                        {
                                            MessageBox.Show("Error exporting output table to CSV file " + strOutputFile);
                                            myFileFuncs.WriteLine(strLogFile, "Error exporting output table to CSV file " + strOutputFile);
                                            this.Cursor = Cursors.Default;
                                            myArcMapFuncs.ToggleDrawing();
                                            myArcMapFuncs.ToggleTOC();
                                            lblPartner.Text = "";
                                            this.BringToFront();
                                            this.Cursor = Cursors.Default;
                                            return;
                                        }
                                        blAppend = true;
                                    }
                                    // Also export the second table - append if necessary.
                                    if (intPolyCount > 0)
                                    {
                                        myFileFuncs.WriteLine(strLogFile, "Appending polygon results to CSV file " + strOutputFile);
                                        blResult = myArcMapFuncs.CopyToCSV(strOutLayerPoly, strOutputFile, true, blAppend, true);
                                        if (!blResult)
                                        {
                                            MessageBox.Show("Error appending output table to CSV file " + strOutputFile);
                                            myFileFuncs.WriteLine(strLogFile, "Error appending output table to CSV file " + strOutputFile);
                                            this.Cursor = Cursors.Default;
                                            myArcMapFuncs.ToggleDrawing();
                                            myArcMapFuncs.ToggleTOC();
                                            lblPartner.Text = "";
                                            this.BringToFront();
                                            this.Cursor = Cursors.Default;
                                            return;
                                        }
                                    }
                                }
                                else
                                {
                                    string strOutputFile = myConfig.GetDefaultPath() + @"\" + strShortName + @"\" + strSQLLayer + ".csv";
                                    myFileFuncs.WriteLine(strLogFile, "Exporting CSV file to " + strOutputFile);
                                    blResult = myArcMapFuncs.CopyToCSV(strOutLayer, strOutputFile, false, false);
                                    if (!blResult)
                                    {
                                        MessageBox.Show("Error exporting output table to CSV file " + strOutputFile);
                                        myFileFuncs.WriteLine(strLogFile, "Error exporting output table to CSV file " + strOutputFile);
                                        this.Cursor = Cursors.Default;
                                        myArcMapFuncs.ToggleDrawing();
                                        myArcMapFuncs.ToggleTOC();
                                        lblPartner.Text = "";
                                        this.BringToFront();
                                        this.Cursor = Cursors.Default;
                                        return;
                                    }
                                }

                            }
                            else if (strExport.ToLower() == "txt")
                            {

                                if (blSpatial)
                                {

                                    bool blAppend = false;
                                    string strOutputFile = strDefaultPath + @"\" + strShortName + @"\" + strSQLLayer + ".txt";
                                    if (intPointCount > 0)
                                    {
                                        myFileFuncs.WriteLine(strLogFile, "Copying point results to tab delimited file: " + strOutputFile);
                                        blResult = myArcMapFuncs.CopyToTabDelimitedFile(strOutLayerPoint, strOutputFile, true, blAppend, true);
                                        if (!blResult)
                                        {
                                            MessageBox.Show("Error exporting output table to txt file " + strOutputFile);
                                            myFileFuncs.WriteLine(strLogFile, "Error exporting output table to txt file " + strOutputFile);
                                            this.Cursor = Cursors.Default;
                                            myArcMapFuncs.ToggleDrawing();
                                            myArcMapFuncs.ToggleTOC();
                                            lblPartner.Text = "";
                                            this.BringToFront();
                                            this.Cursor = Cursors.Default;
                                            return;
                                        }
                                        blAppend = true;
                                    }
                                    // Also export the second table - append if necessary.
                                    if (intPolyCount > 0)
                                    {
                                        myFileFuncs.WriteLine(strLogFile, "Appending polygon results to tab delimited file " + strOutputFile);
                                        blResult = myArcMapFuncs.CopyToTabDelimitedFile(strOutLayerPoly, strOutputFile, true, blAppend, true);
                                        if (!blResult)
                                        {
                                            MessageBox.Show("Error appending output table to txt file " + strOutputFile);
                                            myFileFuncs.WriteLine(strLogFile, "Error appending output table to txt file " + strOutputFile);
                                            this.Cursor = Cursors.Default;
                                            myArcMapFuncs.ToggleDrawing();
                                            myArcMapFuncs.ToggleTOC();
                                            lblPartner.Text = "";
                                            this.BringToFront();
                                            this.Cursor = Cursors.Default;
                                            return;
                                        }
                                    }
                                }
                                else
                                {
                                    string strOutputFile = myConfig.GetDefaultPath() + @"\" + strShortName + @"\" + strSQLLayer + ".txt";
                                    myFileFuncs.WriteLine(strLogFile, "Exporting TXT file to " + strOutputFile);
                                    blResult = myArcMapFuncs.CopyToTabDelimitedFile(strOutLayer, strOutputFile, false, false);
                                    if (!blResult)
                                    {
                                        MessageBox.Show("Error exporting output table to TXT file " + strOutputFile);
                                        myFileFuncs.WriteLine(strLogFile, "Error exporting output table to TXT file " + strOutputFile);
                                        this.Cursor = Cursors.Default;
                                        myArcMapFuncs.ToggleDrawing();
                                        myArcMapFuncs.ToggleTOC();
                                        lblPartner.Text = "";
                                        this.BringToFront();
                                        this.Cursor = Cursors.Default;
                                        return;
                                    }
                                }

                            }


                            lblPartner.Text = strPartner + ": Deleting temporary subset tables";
                            this.BringToFront();

                            // Delete the temporary tables in the SQL database
                            strStoredProcedure = "HLClearSppSubset";
                            SqlCommand myCommand2 = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strStoredProcedure, CommandType.StoredProcedure); // Note pass connection by ref here.
                            mySQLServerFuncs.AddSQLParameter(ref myCommand2, "Schema", strDatabaseSchema);
                            mySQLServerFuncs.AddSQLParameter(ref myCommand2, "SpeciesTable", strChosenSQLLayer);
                            mySQLServerFuncs.AddSQLParameter(ref myCommand2, "UserId", strUserID);
                            try
                            {
                                //myFileFuncs.WriteLine(strLogFile, "Opening SQL connection");
                                dbConn.Open();
                                myFileFuncs.WriteLine(strLogFile, "Deleting temporary tables.");
                                string strRowsAffect = myCommand2.ExecuteNonQuery().ToString();
                                //myFileFuncs.WriteLine(strLogFile, "Closing SQL connection");
                                dbConn.Close();
                                myCommand2 = null;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Could not execute stored procedure. System returned the following message: " +
                                    ex.Message);
                                myFileFuncs.WriteLine(strLogFile, "Could not execute stored procedure. System returned the following message: " +
                                    ex.Message);
                                myArcMapFuncs.ToggleDrawing();
                                myArcMapFuncs.ToggleTOC();
                                lblPartner.Text = "";
                                this.BringToFront();
                                this.Cursor = Cursors.Default;
                                return;
                            }

                            myFileFuncs.WriteLine(strLogFile, "Extracted " + strSQLLayer + " from " + strOutputTable + " for partner " + strPartner + ".");

                        }
                        
                        b++; // Next index.
                    }
                }

                lblPartner.Text = "";
                lblPartner.Refresh();
                this.BringToFront();

                // Delete the final temporary spatial table.
                if (strChosenSQLLayer != "")
                {
                    string strSP = "HLClearSpatialSubset";
                    SqlCommand myCommand3 = mySQLServerFuncs.CreateSQLCommand(ref dbConn, strSP, CommandType.StoredProcedure); // Note pass connection by ref here.
                    mySQLServerFuncs.AddSQLParameter(ref myCommand3, "Schema", strDatabaseSchema);
                    mySQLServerFuncs.AddSQLParameter(ref myCommand3, "SpeciesTable", strChosenSQLLayer); 
                    mySQLServerFuncs.AddSQLParameter(ref myCommand3, "UserId", strUserID);
                    try
                    {
                        //myFileFuncs.WriteLine(strLogFile, "Opening SQL connection");
                        dbConn.Open();
                        myFileFuncs.WriteLine(strLogFile, "Deleting temporary spatial tables.");
                        string strRowsAffect = myCommand3.ExecuteNonQuery().ToString();
                        //myFileFuncs.WriteLine(strLogFile, "Closing SQL connection");
                        dbConn.Close();
                        myCommand3 = null;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Could not execute stored procedure. System returned the following message: " +
                            ex.Message);
                        myFileFuncs.WriteLine(strLogFile, "Could not execute stored procedure. System returned the following message: " +
                            ex.Message);
                        myArcMapFuncs.ToggleDrawing();
                        myArcMapFuncs.ToggleTOC();
                        lblPartner.Text = "";
                        this.BringToFront();
                        this.Cursor = Cursors.Default;
                        return;
                    }
                }


                // Now let's do the GIS layers.
                myFileFuncs.WriteLine(strLogFile, "");
                myFileFuncs.WriteLine(strLogFile, "Processing GIS layers for partner " + strPartner + ".");

                int intLayerIndex = 0;
                foreach (string strGISLayer in liChosenGISLayers)
                {
                    string strChosenFile = liChosenGISFiles[intLayerIndex];
                    // Does the partner want this layer?
                    if (liMapFiles.Contains(strChosenFile))
                    {
                        lblPartner.Text = strPartner + ": Processing GIS layer " + strChosenFile + ".";
                        lblPartner.Refresh();
                        this.BringToFront();

                        // If so, process. 
                        //// Find out all the info for this layer - firstly, in which extracts is it used? NOTE this could be included if people want
                        // more than one extract from a layer. Commented out for the moment and would need reimplementing.
                        //List<int> liMapLayerIndices = new List<int>();
                        //int intIndex = 0;
                        //foreach (string strMapName in liChosenGISLayers)
                        //{
                        //    if (strMapName == strGISLayer)
                        //        liMapLayerIndices.Add(intIndex);
                        //    intIndex++;
                        //}
                        
                        // Now for all these indices process the data.
                        //int intLayerIndex = liMapLayers.IndexOf(strGISLayer);
                        //foreach (int intLayerIndex in liMapLayerIndices)
                        //{
                        string strLayerName = strGISLayer;
                        string strOutputName = strGISLayer; // Input and output are the same name. //liChosenGISFiles[intLayerIndex];
                        string strLayerColumns = liChosenGISColumns[intLayerIndex];
                        string strLayerClause = liChosenGISClauses[intLayerIndex];

                        List<string> liFields = new List<string>();
                        List<string> liRawFields = strLayerColumns.Split(',').ToList();
                        foreach (string strField in liRawFields)
                        {
                            liFields.Add(strField.Trim());
                        }

                        myFileFuncs.WriteLine(strLogFile, "Processing layer " + strLayerName + " to create the output layer " + strOutputName + ".");
                        myFileFuncs.WriteLine(strLogFile, "Selecting features in layer " + strLayerName + " that intersect partner boundary.");
                        // Firstly do the spatial selection.
                        myArcMapFuncs.SelectLayerByLocation(strLayerName, strPartnerTable);

                        int intSelectedFeatures = myArcMapFuncs.CountSelectedLayerFeatures(strLayerName); // How many features are selected?

                        // Now do the attribute selection if required.
                        if (strLayerClause != "" && intSelectedFeatures > 0) 
                        {
                            myFileFuncs.WriteLine(strLogFile, "Creating subset selection for layer " + strLayerName + " using attribute query:");
                            myFileFuncs.WriteLine(strLogFile, strLayerClause);
                            myArcMapFuncs.SelectLayerByAttributes(strLayerName, strLayerClause, "SUBSET_SELECTION");
                            intSelectedFeatures = myArcMapFuncs.CountSelectedLayerFeatures(strLayerName); // How many features are now selected?
                        }

                        if (intSelectedFeatures > 0) // If there is a selection, process.
                        {
                            myFileFuncs.WriteLine(strLogFile, "There are " + intSelectedFeatures.ToString() + " features selected in " + strGISLayer + ".");

                            lblPartner.Text = strPartner + ": Exporting selection in layer " + strChosenFile + ".";
                            lblPartner.Refresh();
                            this.BringToFront();

                            // Export to Shapefile. Create new folder if required.
                            string strOutFolder = strDefaultPath + @"\" + strShortName;

                            if (!myFileFuncs.DirExists(strOutFolder))
                            {
                                myFileFuncs.WriteLine(strLogFile, "Output folder " + strOutFolder + " doesn't exist. Creating...");
                                try
                                {
                                    Directory.CreateDirectory(strOutFolder);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Cannot create directory " + strOutFolder + ". System error: " + ex.Message);
                                    myFileFuncs.WriteLine(strLogFile, "Cannot create directory " + strOutFolder + ". System error: " + ex.Message);
                                    myArcMapFuncs.ToggleDrawing();
                                    myArcMapFuncs.ToggleTOC();
                                    lblPartner.Text = "";
                                    this.BringToFront();
                                    this.Cursor = Cursors.Default;
                                    return;
                                }
                                myFileFuncs.WriteLine(strLogFile, "Output folder created.");
                            }

                            // Create Geodatabase if required.
                            bool blIsGDB = false;
                            if (strFormat == "GDB")
                            {
                                blIsGDB = true;
                                strOutFolder = strOutFolder + "\\DataExtracts.gdb";
                                if (!myFileFuncs.DirExists(strOutFolder))
                                {
                                    myFileFuncs.WriteLine(strLogFile, "Output geodatabase " + strOutFolder + " doesn't exist. Creating...");
                                    myArcMapFuncs.CreateGeodatabase(strOutFolder);
                                }
                                myFileFuncs.WriteLine(strLogFile, "Output geodatabase created.");
                            }


                            // Copy features
                            string strOutLayer = strOutFolder + @"\" + strOutputName; //strLayerName;
                            if (!blIsGDB) strOutLayer = strOutLayer + ".shp"; // While ArcGIS does this automatically it's more clear for the log.
                            myFileFuncs.WriteLine(strLogFile, "Exporting selection to GIS file: " + strOutLayer + ".");
                            myArcMapFuncs.CopyFeatures(strLayerName, strOutLayer); // Copies both to GDB and shapefile.

                            // Drop non-required fields.
                            myFileFuncs.WriteLine(strLogFile, "Removing non-required fields from the output file.");
                            myArcMapFuncs.KeepSelectedFields(strOutLayer, liFields);
                            myArcMapFuncs.RemoveLayer(strOutputName);


                            // Export to CSV if requested as well.
                            lblPartner.Text = strPartner + ": Exporting layer " + strChosenFile + " to text.";
                            lblPartner.Refresh();
                            this.BringToFront();

                            strExport = strExport.ToLower().Trim();
                            if (strExport == "csv")
                            {
                                string strOutTable = myConfig.GetDefaultPath() + @"\" + strShortName + @"\" + strOutputName + ".csv";
                                myFileFuncs.WriteLine(strLogFile, "Exporting CSV file to " + strOutTable);
                                myArcMapFuncs.CopyToCSV(strOutLayer, strOutTable, true, false);
                            }
                            else if (strExport.ToLower() == "txt")
                            {
                                string strOutTable = myConfig.GetDefaultPath() + @"\" + strShortName + @"\" + strOutputName + ".txt";
                                myFileFuncs.WriteLine(strLogFile, "Exporting TXT file to " + strOutTable);
                                myArcMapFuncs.CopyToTabDelimitedFile(strOutLayer, strOutTable, true, false);
                            }

                            myFileFuncs.WriteLine(strLogFile, "Extracted " + strOutputName + " from " + strGISLayer + " for partner " + strPartner + ".");
                            // Clear selected features
                            myArcMapFuncs.ClearSelection(strLayerName);

                        }
                        else
                        {
                            myFileFuncs.WriteLine(strLogFile, "There are no features selected in " + strGISLayer + ".");
                        }
                        //myFileFuncs.WriteLine(strLogFile, "Process complete for GIS layer " + strGISLayer + ", extracting layer " + strOutputName + " for partner " + strPartner + ".");
                        //}
                    }
                    myFileFuncs.WriteLine(strLogFile, "Process complete for GIS layer " + strGISLayer + " for partner " + strPartner + ".");
                    intLayerIndex++;
                }


                // Zip up the results if required
                blCreateZip = false; // For the moment this part of the functionality is disabled.
                if (blCreateZip)
                {
                    lblPartner.Text = strPartner + ": Creating zip file.";
                    lblPartner.Refresh();
                    this.BringToFront();
                    string strSourcePath = strDefaultPath + @"\" + strShortName;
                    string strOutPath = strDefaultPath + @"\" + strShortName + ".zip";

                    // If a previous zip file exists, delete it.
                    string strPreviousZip = strDefaultPath + @"\" + strShortName + @"\" + strShortName + ".zip";
                    if (myFileFuncs.FileExists(strPreviousZip))
                        myFileFuncs.DeleteFile(strPreviousZip);

                    if (myFileFuncs.DirExists(strSourcePath))
                    {
                        
                        myFileFuncs.WriteLine(strLogFile, "Creating zip file " + strOutPath + "......");

                        ArchiveCreator AC = new WinBaseZipCreator(strOutPath);
                        ArchiveObj anObj = AC.GetArchive();
                        anObj.AddAllFiles(strSourcePath);

                        myFileFuncs.WriteLine(strLogFile, "Writing zip file...");
                        try
                        {
                            anObj.SaveArchive();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Could not write zip file for partner " + strPartner + ". Error message is " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            myFileFuncs.WriteLine(strLogFile, "Could not write zip file. Error message is " + ex.Message);
                            myFileFuncs.WriteLine(strLogFile, "Zip file not created for partner " + strPartner);
                            anObj = null;
                            AC = null;
                            if (myFileFuncs.FileExists(strDefaultPath + @"\" + strShortName + ".zip"))
                                myFileFuncs.DeleteFile(strDefaultPath + @"\" + strShortName + ".zip");
                        }
                        anObj = null;
                        AC = null;
                        // Move the zip file to its final location
                        myFileFuncs.WriteLine(strLogFile, "Moving zip file to final location: " + strOutPath + ".");
                        string strFinalOutPath = strDefaultPath + @"\" + strShortName + @"\" + strShortName + ".zip";
                        if (myFileFuncs.FileExists(strFinalOutPath))
                            myFileFuncs.DeleteFile(strFinalOutPath);
                        File.Move(strOutPath, strFinalOutPath);
                        myFileFuncs.WriteLine(strLogFile, "Zip file created successfully for partner " + strPartner + ".");
                        
                        
                    }
                }
                
                // Log the completion of this partner.
                myFileFuncs.WriteLine(strLogFile, "Process complete for partner " + strPartner);
            }

            // Clear the selection on the partner layer
            myArcMapFuncs.ClearSelection(strPartnerTable);

            lblPartner.Text = "";
            lblPartner.Refresh();
            this.BringToFront();

            // Switch drawing back on.
            myArcMapFuncs.ToggleDrawing();
            myArcMapFuncs.ToggleTOC();
            this.BringToFront();
            this.Cursor = Cursors.Default;

            myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");
            myFileFuncs.WriteLine(strLogFile, "Process complete");
            myFileFuncs.WriteLine(strLogFile, "-----------------------------------------------------------------------");

            DialogResult dlResult = MessageBox.Show("Process complete. Do you wish to close the form?", "Data Extractor", MessageBoxButtons.YesNo);
            if (dlResult == System.Windows.Forms.DialogResult.Yes)
                this.Close();
            else this.BringToFront();

            Process.Start("notepad.exe", strLogFile);
            return;
        }

        private void frmDataExtractor_Load(object sender, EventArgs e)
        {

        }

        private void lstActivePartners_SelectedIndexChanged(object sender, EventArgs e)
        {
            // This is not a needed event
        }

        private void lstActivePartners_DoubleClick(object sender, EventArgs e)
        {
            // Show the comment that is related to the selected item.
            if (lstActivePartners.SelectedItem != null)
            {
                // Get the partner name
                string strPartner = lstActivePartners.SelectedItem.ToString();
                // Build the query
                string strQuery = "SELECT " + myConfig.GetNotesColumn() + " FROM " + myConfig.GetPartnerTable() + " WHERE " +
                                  myConfig.GetPartnerColumn() + " = '" + strPartner + "'";
                SqlConnection myConn = mySQLServerFuncs.CreateSQLConnection(myConfig.GetConnectionString());
                myConn.Open();
                string strNote = mySQLServerFuncs.GetStringValue(ref myConn, strQuery);
                MessageBox.Show(strNote, "Partner Note", MessageBoxButtons.OK, MessageBoxIcon.Information);
                myConn.Close();
            }
        }
    }
}
