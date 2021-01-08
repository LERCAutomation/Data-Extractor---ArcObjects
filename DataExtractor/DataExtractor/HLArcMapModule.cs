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
using System.IO;
using System.Windows.Forms;
using System.Threading;

using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Desktop.AddIns;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.GeoDatabaseUI;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.DataSourcesOleDB;

using ESRI.ArcGIS.Catalog;
using ESRI.ArcGIS.CatalogUI;
using ESRI.ArcGIS.Display;

using HLFileFunctions;

namespace HLArcMapModule
{
    class ArcMapFunctions
    {
        #region Constructor
        private IApplication thisApplication;
        private FileFunctions myFileFuncs;
        // Class constructor.
        public ArcMapFunctions(IApplication theApplication)
        {
            // Set the application for the class to work with.
            // Note the application can be got at from a command / tool by using
            // IApplication pApp = ArcMap.Application - then pass pApp as an argument.
            this.thisApplication = theApplication;
            myFileFuncs = new FileFunctions();
        }
        #endregion

        public IMxDocument GetIMXDocument()
        {
            ESRI.ArcGIS.ArcMapUI.IMxDocument mxDocument = ((ESRI.ArcGIS.ArcMapUI.IMxDocument)(thisApplication.Document));
            return mxDocument;
        }

        public ESRI.ArcGIS.Carto.IMap GetMap()
        {
            if (thisApplication == null)
            {
                return null;
            }
            ESRI.ArcGIS.ArcMapUI.IMxDocument mxDocument = ((ESRI.ArcGIS.ArcMapUI.IMxDocument)(thisApplication.Document)); // Explicit Cast
            ESRI.ArcGIS.Carto.IActiveView activeView = mxDocument.ActiveView;
            ESRI.ArcGIS.Carto.IMap map = activeView.FocusMap;

            return map;
        }

        public void RefreshTOC()
        {
            IMxDocument theDoc = GetIMXDocument();
            theDoc.CurrentContentsView.Refresh(null);
        }

        public IWorkspaceFactory GetWorkspaceFactory(string aFilePath, bool aTextFile = false, bool Messages = false)
        {
            // This function decides what type of feature workspace factory would be best for this file.
            // it is up to the user to decide whether the file path and file names exist (or should exist).

            // Reworked 18/05/2016 to deal with the singleton issue.

            IWorkspaceFactory pWSF;
            // What type of output file it it? This defines what kind of workspace factory will be returned.
            if (aFilePath.Substring(aFilePath.Length - 4, 4).ToLower() == ".gdb")
            {
                // It is a file geodatabase file.
                Type t = Type.GetTypeFromProgID("esriDataSourcesGDB.FileGDBWorkspaceFactory");
                System.Object obj = Activator.CreateInstance(t);
                pWSF = obj as IWorkspaceFactory;
            }
            else if (aFilePath.Substring(aFilePath.Length - 4, 4).ToLower() == ".mdb")
            {
                // Personal geodatabase.
                Type t = Type.GetTypeFromProgID("esriDataSourcesGDB.AccessWorkspaceFactory");
                System.Object obj = Activator.CreateInstance(t);
                pWSF = obj as IWorkspaceFactory;
            }
            else if (aFilePath.Substring(aFilePath.Length - 4, 4).ToLower() == ".sde")
            {
                // ArcSDE connection
                Type t = Type.GetTypeFromProgID("esriDataSourcesGDB.SdeWorkspaceFactory");
                System.Object obj = Activator.CreateInstance(t);
                pWSF = obj as IWorkspaceFactory;
            }
            else if (aTextFile == true)
            {
                // Text file
                //Type t = Type.GetTypeFromProgID("esriDataSourcesOleDB.TextFileWorkspaceFactory");
                //System.Object obj = Activator.CreateInstance(t);
                pWSF = new TextFileWorkspaceFactory();
            }
            else // Shapefile
            {
                Type t = Type.GetTypeFromProgID("esriDataSourcesFile.ShapefileWorkspaceFactory");
                System.Object obj = Activator.CreateInstance(t);
                pWSF = obj as IWorkspaceFactory;
            }
            return pWSF;
        }


        public bool CreateGeodatabase(string aFilePath, bool Messages = false)
        {
            // Check the file name given.
            if (aFilePath.Substring(aFilePath.Length - 4, 4).ToLower() != ".gdb")
            {
                if (Messages) MessageBox.Show("The file path " + aFilePath + " is not a valid geodatabase name");
                return false;
            }

            if (myFileFuncs.DirExists(aFilePath))
            {
                if (Messages) MessageBox.Show("The geodatabase " + aFilePath + " already exists. Not created");
                return false;
            }
            try
            {
                Type factoryType = Type.GetTypeFromProgID("esriDataSourcesGDB.FileGDBWorkspaceFactory");
                IWorkspaceFactory WorkspaceFactory = (IWorkspaceFactory)Activator.CreateInstance(factoryType);
                IWorkspaceName WorkspaceName = WorkspaceFactory.Create(myFileFuncs.GetDirectoryName(aFilePath), myFileFuncs.GetFileName(aFilePath), null, thisApplication.hWnd);
            }
            catch (Exception ex)
            {
                if (Messages) MessageBox.Show("Could not create file gdb " + aFilePath + ". Error message: " + ex.Message);
                return false;
            }
            

            return true;
        }

        #region FeatureclassExists
        public bool FeatureclassExists(string aFilePath, string aDatasetName)
        {
            
            if (aDatasetName.Substring(aDatasetName.Length - 4, 1) == ".")
            {
                // it's a file.
                if (myFileFuncs.FileExists(aFilePath + @"\" + aDatasetName))
                    return true;
                else
                    return false;
            }
            else if (aFilePath.Substring(aFilePath.Length - 3, 3).ToLower() == "sde")
            {
                // It's an SDE class
                // Not handled. We know the table exists.
                return true;
            }
            else // it is a geodatabase class.
            {
                IWorkspaceFactory pWSF = GetWorkspaceFactory(aFilePath);
                IWorkspace2 pWS = (IWorkspace2)pWSF.OpenFromFile(aFilePath, 0);
                if (pWS.get_NameExists(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTFeatureClass, aDatasetName))
                    return true;
                else
                    return false;
            }
        }

        public bool FeatureclassExists(string aFullPath)
        {
            return FeatureclassExists(myFileFuncs.GetDirectoryName(aFullPath), myFileFuncs.GetFileName(aFullPath));
        }
        #endregion

        #region GetFeatureClass
        public IFeatureClass GetFeatureClass(string aFilePath, string aDatasetName, bool Messages = false)
        // This is incredibly quick.
        {
            // Check input first.
            string aTestPath = aFilePath;
            if (aFilePath.ToLower().Contains(".sde"))
            {
                aTestPath = myFileFuncs.GetDirectoryName(aFilePath);
            }
            if (myFileFuncs.DirExists(aTestPath) == false || aDatasetName == null)
            {
                if (Messages) MessageBox.Show("Please provide valid input", "Get Featureclass");
                return null;
            }
            

            IWorkspaceFactory pWSF = GetWorkspaceFactory(aFilePath);
            IFeatureWorkspace pWS = (IFeatureWorkspace)pWSF.OpenFromFile(aFilePath, 0);
            if (FeatureclassExists(aFilePath, aDatasetName))
            {
                IFeatureClass pFC = pWS.OpenFeatureClass(aDatasetName);
                return pFC;
            }
            else
            {
                if (Messages) MessageBox.Show("The file " + aDatasetName + " doesn't exist in this location", "Open Feature Class from Disk");
                return null;
            }
            
        }

        public IFeatureClass GetFeatureClass(string aFullPath, bool Messages = false)
        {
            string aFilePath = myFileFuncs.GetDirectoryName(aFullPath);
            string aDatasetName = myFileFuncs.GetFileName(aFullPath);
            IFeatureClass pFC = GetFeatureClass(aFilePath, aDatasetName, Messages);
            return pFC;
        }

        #endregion

        public IFeatureLayer GetFeatureLayerFromString(string aFeatureClassName, bool Messages = false)
        {
            // as far as I can see this does not work for geodatabase files.
            // firstly get the Feature Class
            // Does it exist?
            if (!myFileFuncs.FileExists(aFeatureClassName))
            {
                if (Messages)
                {
                    MessageBox.Show("The featureclass " + aFeatureClassName + " does not exist");
                }
                return null;
            }
            string aFilePath = myFileFuncs.GetDirectoryName(aFeatureClassName);
            string aFCName = myFileFuncs.GetFileName(aFeatureClassName);

            IFeatureClass myFC = GetFeatureClass(aFilePath, aFCName);
            if (myFC == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot open featureclass " + aFeatureClassName);
                }
                return null;
            }

            // Now get the Feature Layer from this.
            FeatureLayer pFL = new FeatureLayer();
            pFL.FeatureClass = myFC;
            pFL.Name = myFC.AliasName;
            return pFL;
        }

        public ILayer GetLayer(string aName, bool Messages = false)
        {
            // Gets existing layer in map.
            // Check there is input.
           if (aName == null)
           {
               if (Messages)
               {
                   MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
               }
               return null;
            }
        
            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages)
                {
                    MessageBox.Show("No map found", "Find Layer By Name");
                }
                return null;
            }
            IEnumLayer pLayers = pMap.Layers;
            Boolean blFoundit = false;
            ILayer pTargetLayer = null;

            ILayer pLayer = pLayers.Next();

            // Look through the layers and carry on until found,
            // or we have reached the end of the list.
            while ((pLayer != null) && !blFoundit)
            {
                if (!(pLayer is ICompositeLayer))
                {
                    if (pLayer.Name.Equals(aName, StringComparison.OrdinalIgnoreCase))
                    {
                        pTargetLayer = pLayer;
                        blFoundit = true;
                    }
                }
                pLayer = pLayers.Next();
            }

            if (pTargetLayer == null)
            {
                if (Messages) MessageBox.Show("The layer " + aName + " doesn't exist", "Find Layer");
                return null;
            }
            return pTargetLayer;
        }

        public bool FieldExists(string aFilePath, string aDatasetName, string aFieldName, bool Messages = false)
        {
            // This function returns true if a field (or a field alias) exists, false if it doesn't (or the dataset doesn't)
            IFeatureClass myFC = GetFeatureClass(aFilePath, aDatasetName);
            ITable myTab;
            if (myFC == null)
            {
                myTab = GetTable(aFilePath, aDatasetName);
                if (myTab == null) return false; // Dataset doesn't exist.
            }
            else
            {
                myTab = (ITable)myFC;
            }

            int aTest;
            IFields theFields = myTab.Fields;
            aTest = theFields.FindField(aFieldName);
            if (aTest == -1)
            {
                aTest = theFields.FindFieldByAliasName(aFieldName);
            }

            if (aTest == -1) return false;
            return true;
        }

        public bool FieldExists(ILayer aLayer, string aFieldName, string aLogFile = "", bool Messages = false)
        {
            IFeatureLayer pFL = null;
            try
            {
                pFL = (IFeatureLayer)aLayer;
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("The layer is not a feature layer");
                if (aLogFile != "")
                    myFileFuncs.WriteLine(aLogFile, "Function FieldExists returned the following error: The input layer aLayer is not a feature layer.");
                return false;
            }
            IFeatureClass pFC = pFL.FeatureClass;
            return FieldExists(pFC, aFieldName);
        }

        public bool FieldExists(IFeatureClass aFeatureClass, string aFieldName, string aLogFile = "", bool Messages = false)
        {

            //int aTest;
            IFields theFields = aFeatureClass.Fields;
            return FieldExists(theFields, aFieldName, aLogFile, Messages);
            //aTest = theFields.FindField(aFieldName);
            //if (aTest == -1)
            //{
            //    aTest = theFields.FindFieldByAliasName(aFieldName);
            //}

            //if (aTest == -1) return false;
            //return true;
        }

        public bool FieldExists(IFields theFields, string aFieldName, string aLogFile = "", bool Messages = false)
        {
            int aTest;
            aTest = theFields.FindField(aFieldName);
            if (aTest == -1)
                aTest = theFields.FindFieldByAliasName(aFieldName);
            if (aTest == -1) return false;
            return true;
        }

        public bool KeepSelectedFields(string aFeatureClassOrLayer, List<string> aFieldList, bool Messages = false)
        {
            // This function deletes all the fields from aFeatureClassOrLayer that are not required and are not in aFieldList.
            IFeatureClass pFC = null;
            if (LayerExists(aFeatureClassOrLayer))
            {
                ILayer pLayer = GetLayer(aFeatureClassOrLayer);
                if (pLayer is IFeatureLayer)
                {
                    IFeatureLayer pFL = (IFeatureLayer)pLayer;
                }
                else
                {
                    if (Messages) MessageBox.Show("The layer " + aFeatureClassOrLayer + " is not a feature layer.");
                    return false;
                }

            }
            else
            {
                pFC = GetFeatureClass(aFeatureClassOrLayer);
                if (pFC == null)
                {
                    if (Messages) MessageBox.Show("The feature class " + aFeatureClassOrLayer + " doesn't exist.");
                    return false;
                }
            }

            // Now drop any fields from the output that we don't want.
            IFields pFields = pFC.Fields;
            List<string> strDeleteFields = new List<string>();

            // Make a list of fields to delete.
            for (int i = 0; i < pFields.FieldCount; i++)
            {
                IField pField = pFields.get_Field(i);
                // Does it exist in the 'keep' list or is it required?
                if (aFieldList.IndexOf(pField.Name) == -1 && !pField.Required)
                {
                    // If not, add to the delete list.
                    strDeleteFields.Add(pField.Name);
                }
            }

            //Delete the listed fields.
            foreach (string strField in strDeleteFields)
            {
                DeleteField(pFC, strField);
            }
            return true;
        }

        public bool KeepSelectedTableFields(string aTableName, List<string> aFieldList, bool Messages = false)
        {
            // firstly get the Table
            // Does it exist? // Does not work for GeoDB tables!!
            if (!myFileFuncs.FileExists(aTableName))
            {
                if (Messages)
                {
                    MessageBox.Show("The table " + aTableName + " does not exist");
                }
                return false;
            }
            string aFilePath = myFileFuncs.GetDirectoryName(aTableName);
            string aTabName = myFileFuncs.GetFileName(aTableName);

            ITable myTab = null;
            myTab = GetTable(aFilePath, aTabName);
            if (myTab == null)
            {
                if (Messages) MessageBox.Show("The table " + aTableName + " doesn't exist.");
                return false; // Dataset doesn't exist.
            }

            // Now drop any fields from the output that we don't want.
            IFields pFields = myTab.Fields;
            List<string> strDeleteFields = new List<string>();

            // Make a list of fields to delete.
            for (int i = 0; i < pFields.FieldCount; i++)
            {
                IField pField = pFields.get_Field(i);
                if (aFieldList.IndexOf(pField.Name) == -1 && !pField.Required)
                // Does it exist in the 'keep' list or is it required?
                {
                    // If not, add to the delete list.
                    strDeleteFields.Add(pField.Name);
                }
            }

            //Delete the listed fields.
            foreach (string strField in strDeleteFields)
            {
                DeleteField(myTab, strField);
            }
            return true;
        }

        public bool DeleteField(IFeatureClass aFeatureClass, string aFieldName)
        {
            // Get the fields collection
            int intIndex = aFeatureClass.Fields.FindField(aFieldName);
            IField pField = aFeatureClass.Fields.get_Field(intIndex);
            try
            {
                aFeatureClass.DeleteField(pField);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot delete field " + aFieldName + ". System error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public bool DeleteField(ITable aTable, string aFieldName)
        {
            // Get the fields collection
            int intIndex = aTable.Fields.FindField(aFieldName);
            IField pField = aTable.Fields.get_Field(intIndex);
            try
            {
                aTable.DeleteField(pField);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot delete field " + aFieldName + ". System error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public List<string> FieldsExist(string aFilePath, string aDatasetName, List<string> FieldNames, bool Messages = false)
        {
            List<string> FieldsThatExist = new List<string>();
            foreach (string aFieldName in FieldNames)
            {
                if (FieldExists(aFilePath, aDatasetName, aFieldName))
                    FieldsThatExist.Add(aFieldName);
            }
            return FieldsThatExist;
        }

        public List<string> FieldsExist(string aName, List<string> FieldNames, bool Messages = false)
        {
            IFeatureLayer featureLayer = null;
            try
            {
                featureLayer = (IFeatureLayer)GetLayer(aName);

                List<string> FieldsThatExist = new List<string>();
                foreach (string aFieldName in FieldNames)
                {
                    if (FieldExists(featureLayer, aFieldName))
                        FieldsThatExist.Add(aFieldName);
                }
                return FieldsThatExist;
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("The layer " + aName + " is not a feature layer");
                return new List<string>();
            }
        }

        /// <summary>
        /// Calculate the total row length for a table.
        /// </summary>
        /// <param name="aName"></param>
        /// <param name="Messages"></param>
        /// <returns></returns>
        public int GetRowLength(string aFilePath, string aDatasetName, bool Messages = false)
        {
            ITable myTab = GetTable(aFilePath, aDatasetName);
            if (myTab == null) return 0; // Dataset doesn't exist.

            IFields theFields = myTab.Fields;
            int intFieldCount = theFields.FieldCount;
            int intRowLength = 1;

            // iterate through the fields in the collection.
            for (int i = 0; i < intFieldCount; i++)
            {
                // Get the field at the given index.
                IField pField = theFields.get_Field(i);
                string strField = pField.Name;
                string strType = pField.Type.ToString();

                int intFieldLen;
                if (strType == "esriFieldTypeInteger")
                    intFieldLen = 10;
                else if (strType == "esriFieldTypeGeometry")
                    intFieldLen = 0;
                else
                    intFieldLen = pField.Length;
                intRowLength = intRowLength + intFieldLen;
            }

            return intRowLength;
        }

        public bool AddLayerFromFClass(IFeatureClass theFeatureClass, bool Messages = false)
        {
            // Check we have input
            if (theFeatureClass == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Please pass a feature class", "Add Layer From Feature Class");
                }
                return false;
            }
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages)
                {
                    MessageBox.Show("No map found", "Add Layer From Feature Class");
                }
                return false;
            }
            FeatureLayer pFL = new FeatureLayer();
            pFL.FeatureClass = theFeatureClass;
            List<string> strNameParts = theFeatureClass.AliasName.Split('.').ToList();
            pFL.Name = strNameParts[strNameParts.Count - 1]; // take the last entry.
            //pFL.Name = theFeatureClass.AliasName;
            pMap.AddLayer(pFL);

            return true;
        }

        public bool AddFeatureLayerFromString(string aFeatureClassName, bool Messages = false)
        {
            // firstly get the Feature Class
            // Does it exist?
            if (!myFileFuncs.FileExists(aFeatureClassName) && !FeatureclassExists(aFeatureClassName))
            {
                if (Messages)
                {
                    MessageBox.Show("The featureclass " + aFeatureClassName + " does not exist");
                }
                return false;
            }
            string aFilePath = myFileFuncs.GetDirectoryName(aFeatureClassName);
            string aFCName = myFileFuncs.GetFileName(aFeatureClassName);

            IFeatureClass myFC = GetFeatureClass(aFilePath, aFCName);
            if (myFC == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot open featureclass " + aFeatureClassName);
                }
                return false;
            }

            // Now add it to the view.
            bool blResult = AddLayerFromFClass(myFC);
            if (blResult)
            {
                return true;
            }
            else
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot add featureclass " + aFeatureClassName);
                }
                return false;
            }
        }

        #region TableExists
        public bool TableExists(string aFilePath, string aDatasetName)
        {

            if (aDatasetName.Substring(aDatasetName.Length - 4, 1) == ".")
            {
                // it's a file.
                if (myFileFuncs.FileExists(aFilePath + @"\" + aDatasetName))
                    return true;
                else
                    return false;
            }
            else if (aFilePath.Substring(aFilePath.Length - 3, 3).ToLower() == "sde")
            {
                // It's an SDE class
                // Not handled. We know the table exists.
                return true;
            }
            else // it is a geodatabase class.
            {
                IWorkspaceFactory pWSF = GetWorkspaceFactory(aFilePath);
                IWorkspace2 pWS = (IWorkspace2)pWSF.OpenFromFile(aFilePath, 0);
                if (pWS.get_NameExists(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTTable, aDatasetName))
                    return true;
                else
                    return false;
            }
        }

        public bool TableExists(string aFullPath)
        {
            return TableExists(myFileFuncs.GetDirectoryName(aFullPath), myFileFuncs.GetFileName(aFullPath));
        }
        #endregion

        #region GetTable
        public ITable GetTable(string aFilePath, string aDatasetName, bool Messages = false)
        {
            // Check input first.
            string aTestPath = aFilePath;
            if (aFilePath.ToLower().Contains(".sde"))
            {
                aTestPath = myFileFuncs.GetDirectoryName(aFilePath);
            }
            if (myFileFuncs.DirExists(aTestPath) == false || aDatasetName == null)
            {
                if (Messages) MessageBox.Show("Please provide valid input", "Get Table");
                return null;
            }
            bool blText = false;
            string strExt = aDatasetName.Substring(aDatasetName.Length - 4, 4).ToLower();
            if (strExt.ToLower() == ".txt" || strExt.ToLower() == ".csv" || strExt.ToLower() == ".tab")
            {
                blText = true;
            }

            IWorkspaceFactory pWSF = GetWorkspaceFactory(aFilePath, blText);
            IFeatureWorkspace pWS = (IFeatureWorkspace)pWSF.OpenFromFile(aFilePath, 0);
            ITable pTable = pWS.OpenTable(aDatasetName);
            if (pTable == null)
            {
                if (Messages) MessageBox.Show("The file " + aDatasetName + " doesn't exist in this location", "Open Table from Disk");
                return null;
            }
            return pTable;
        }

        public ITable GetTable(string aTableLayer, bool Messages = false)
        {
            IMap pMap = GetMap();
            ITableCollection pColl = (ITableCollection)pMap;
            IStandaloneTable pThisTable = null;

            for (int I = 0; I < pColl.TableCount; I++)
            {
                pThisTable = (IStandaloneTable)pColl.Table[I];
                if (pThisTable.Name.Equals(aTableLayer, StringComparison.OrdinalIgnoreCase))
                {
                    ITable myTable = pThisTable.Table;
                    return myTable;
                }
            }
            if (Messages)
            {
                MessageBox.Show("The table layer " + aTableLayer + " could not be found in this map");
            }
            return null;
        }
        #endregion

        public ITable GetTable(IWorkspace aWorkspace, string aDatasetName, bool Messages = false)
        {
            IFeatureWorkspace pWS = (IFeatureWorkspace)aWorkspace;
            ITable pTable = pWS.OpenTable(aDatasetName);
            if (pTable == null)
            {
                if (Messages) MessageBox.Show("The file " + aDatasetName + " doesn't exist in this location", "Open Table from Disk");
                return null;
            }
            return pTable;
        }

        public ICursor GetCursorOnFeatureLayer(string aLayerName, bool GetSelected = true, bool Messages = false)
        {
            ILayer pLayer = GetLayer(aLayerName);
            if (pLayer == null)
            {
                if (Messages) MessageBox.Show("The layer " + aLayerName + " is not a layer", "Get Cursor on Feature Layer");
                return null;
            }
            ICursor pCurs = null;
            if (pLayer is IFeatureLayer) // Check that it is a feature layer.
            {
                IFeatureLayer pFL = (IFeatureLayer)pLayer; 
                if (GetSelected)
                {
                    IFeatureSelection pFSel = (IFeatureSelection)pLayer;
                    ISelectionSet pFSelSet = pFSel.SelectionSet;
                    pFSelSet.Search(null, false, out pCurs);
                }
                else
                {
                    IFeatureClass pFC = pFL.FeatureClass;
                    pCurs = (ICursor)pFC.Search(null, false);
                }
            }
            else
            {
                if (Messages) MessageBox.Show("The layer " + aLayerName + " is not a feature layer", "Get Cursor on Feature Layer");
                return null;
            }
            return pCurs;

        }


        public bool AddTableLayerFromString(string aTableName, string aLayerName, bool Messages = false)
        {
            // firstly get the Table
            // Does it exist? // Does not work for GeoDB tables!!
            if (!myFileFuncs.FileExists(aTableName))
            {
                if (Messages)
                {
                    MessageBox.Show("The table " + aTableName + " does not exist");
                }
                return false;
            }
            string aFilePath = myFileFuncs.GetDirectoryName(aTableName);
            string aTabName = myFileFuncs.GetFileName(aTableName);

            ITable myTable = GetTable(aFilePath, aTabName);
            if (myTable == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot open table " + aTableName);
                }
                return false;
            }

            // Now add it to the view.
            bool blResult = AddLayerFromTable(myTable, aLayerName);
            if (blResult)
            {
                return true;
            }
            else
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot add table " + aTabName);
                }
                return false;
            }
        }

        public bool AddLayerFromTable(ITable theTable, string aName, bool Messages = false)
        {
            // check we have nput
            if (theTable == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Please pass a table", "Add Layer From Table");
                }
                return false;
            }
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages)
                {
                    MessageBox.Show("No map found", "Add Layer From Table");
                }
                return false;
            }
            IStandaloneTableCollection pStandaloneTableCollection = (IStandaloneTableCollection)pMap;
            IStandaloneTable pTable = new StandaloneTable();
            IMxDocument mxDoc = GetIMXDocument();

            pTable.Table = theTable;
            pTable.Name = aName;

            // Remove if already exists
            if (TableLayerExists(aName))
                RemoveStandaloneTable(aName);

            mxDoc.UpdateContents();
            
            pStandaloneTableCollection.AddStandaloneTable(pTable);
            mxDoc.UpdateContents();
            return true;
        }

        public bool TableLayerExists(string aLayerName, bool Messages = false)
        {
            // Check there is input.
            if (aLayerName == null)
            {
                if (Messages) MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return false;
            }

            // Get map, and layer names.
            IMxDocument mxDoc = GetIMXDocument();
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return false;
            }

            IStandaloneTableCollection pColl = (IStandaloneTableCollection)pMap;
            IStandaloneTable pThisTable = null;
            for (int I = 0; I < pColl.StandaloneTableCount; I++)
            {
                pThisTable = pColl.StandaloneTable[I];
                if (pThisTable.Name.Equals(aLayerName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                    //pColl.RemoveStandaloneTable(pThisTable);
                   // mxDoc.UpdateContents();
                    //break; // important: get out now, the index is no longer valid
                }
            }
            return false;
        }

        public bool RemoveStandaloneTable(string aTableName, bool Messages = false)
        {
            // Check there is input.
            if (aTableName == null)
            {
                if (Messages) MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return false;
            }

            // Get map, and layer names.
            IMxDocument mxDoc = GetIMXDocument();
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return false;
            }

            IStandaloneTableCollection pColl = (IStandaloneTableCollection)pMap;
            IStandaloneTable pThisTable = null;
            for (int I = 0; I < pColl.StandaloneTableCount; I++)
            {
                pThisTable = pColl.StandaloneTable[I];
                if (pThisTable.Name.Equals(aTableName, StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        pColl.RemoveStandaloneTable(pThisTable);
                        mxDoc.UpdateContents();
                        return true; // important: get out now, the index is no longer valid
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
            }
            return false;
        }


        public bool LayerExists(string aLayerName, bool Messages = false)
        {
            // Check there is input.
            if (aLayerName == null)
            {
                if (Messages) MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return false;
            }

            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return false;
            }
            IEnumLayer pLayers = pMap.Layers;

            ILayer pLayer = pLayers.Next();

            // Look through the layers and carry on until found,
            // or we have reached the end of the list.
            while (pLayer != null)
            {
                if (!(pLayer is IGroupLayer))
                {
                    if (pLayer.Name.Equals(aLayerName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                }
                pLayer = pLayers.Next();
            }
            return false;
        }

        public bool GroupLayerExists(string aGroupLayerName, bool Messages = false)
        {
            // Check there is input.
            if (aGroupLayerName == null)
            {
                if (Messages) MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return false;
            }

            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return false;
            }
            IEnumLayer pLayers = pMap.Layers;

            ILayer pLayer = pLayers.Next();

            // Look through the layers and carry on until found,
            // or we have reached the end of the list.
            while (pLayer != null)
            {
                if (pLayer is IGroupLayer)
                {
                    if (pLayer.Name.Equals(aGroupLayerName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                }
                pLayer = pLayers.Next();
            }
            return false;
        }

        public ILayer GetGroupLayer(string aGroupLayerName, bool Messages = false)
        {
            // Check there is input.
            if (aGroupLayerName == null)
            {
                if (Messages) MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return null;
            }

            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return null;
            }
            IEnumLayer pLayers = pMap.Layers;

            ILayer pLayer = pLayers.Next();

            // Look through the layers and carry on until found,
            // or we have reached the end of the list.
            while (pLayer != null)
            {
                if (pLayer is IGroupLayer)
                {
                    if (pLayer.Name.Equals(aGroupLayerName, StringComparison.OrdinalIgnoreCase))
                    {
                        return pLayer;
                    }

                }
                pLayer = pLayers.Next();
            }
            return null;
        }      
        
        public bool MoveToGroupLayer(string theGroupLayerName, ILayer aLayer,  bool Messages = false)
        {
            bool blExists = false;
            IGroupLayer myGroupLayer = new GroupLayer(); 
            // Does the group layer exist?
            if (GroupLayerExists(theGroupLayerName))
            {
                myGroupLayer = (IGroupLayer)GetGroupLayer(theGroupLayerName);
                blExists = true;
            }
            else
            {
                myGroupLayer.Name = theGroupLayerName;
            }
            string theOldName = aLayer.Name;

            // Remove the original instance, then add it to the group.
            RemoveLayer(aLayer);
            myGroupLayer.Add(aLayer);
            
            if (!blExists)
            {
                // Add the layer to the map.
                IMap pMap = GetMap();
                pMap.AddLayer(myGroupLayer);
            }
            RefreshTOC();
            return true;
        }

        #region RemoveLayer
        public bool RemoveLayer(string aLayerName, bool Messages = false)
        {
            // Check there is input.
            if (aLayerName == null)
            {
                MessageBox.Show("Please pass a valid layer name", "Find Layer By Name");
                return false;
            }

            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Name");
                return false;
            }
            IEnumLayer pLayers = pMap.Layers;

            ILayer pLayer = pLayers.Next();

            // Look through the layers and carry on until found,
            // or we have reached the end of the list.
            while (pLayer != null)
            {
                if (!(pLayer is IGroupLayer))
                {
                    if (pLayer.Name.Equals(aLayerName, StringComparison.OrdinalIgnoreCase))
                    {
                        pMap.DeleteLayer(pLayer);
                        return true;
                    }

                }
                pLayer = pLayers.Next();
            }
            return false;
        }

        public bool RemoveLayer(ILayer aLayer, bool Messages = false)
        {
            // Check there is input.
            if (aLayer == null)
            {
                MessageBox.Show("Please pass a valid layer ", "Remove Layer");
                return false;
            }

            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Remove Layer");
                return false;
            }
            pMap.DeleteLayer(aLayer);
            return true;
        }

        public bool RemoveDataset(string aDatasetPath, string aDatasetSchema, string aDatasetName, bool Messages = false)
        {
            // Check there is input.
            if (aDatasetPath == null || aDatasetSchema == null || aDatasetName == null)
            {
                MessageBox.Show("Please pass valid dataset details", "Find Layer By Dataset");
                return false;
            }

            // Get map, and layer names.
            IMap pMap = GetMap();
            if (pMap == null)
            {
                if (Messages) MessageBox.Show("No map found", "Find Layer By Dataset");
                return false;
            }
            IEnumLayer pLayers = pMap.Layers;

            ILayer pLayer = pLayers.Next();

            // Look through the layers and carry on until found,
            // or we have reached the end of the list.
            while (pLayer != null)
            {
                if (!(pLayer is IGroupLayer))
                {
                    if ((pLayer is IDataset) && !(pLayer is IRasterLayer))
                    {
                        IDataset pDataset = (IDataset)pLayer;
                        IWorkspace pWorkspace = (IWorkspace)pDataset.Workspace;
                        string strDatasetPath = pWorkspace.PathName;
                        string strDatasetName = pDataset.Name;

                        if (strDatasetPath.Equals(aDatasetPath, StringComparison.OrdinalIgnoreCase) &&
                            strDatasetName.EndsWith(aDatasetSchema + "." + aDatasetName, StringComparison.OrdinalIgnoreCase))
                        {
                            pMap.DeleteLayer(pLayer);
                            return true;
                        }

                    }
                }
                pLayer = pLayers.Next();
            }
            return false;
        }
        #endregion


        public bool SelectLayerByAttributes(string aFeatureLayerName, string aWhereClause, string aSelectionType = "NEW_SELECTION", bool Messages = false)
        {
            ///<summary>Select features in the IActiveView by an attribute query using a SQL syntax in a where clause.</summary>
            /// 
            ///<param name="featureLayer">An IFeatureLayer interface to select upon</param>
            ///<param name="whereClause">A System.String that is the SQL where clause syntax to select features. Example: "CityName = 'Redlands'"</param>
            ///  
            ///<remarks>Providing and empty string "" will return all records.</remarks>
            if (!LayerExists(aFeatureLayerName))
                return false;

            IFeatureLayer featureLayer = null;
            try
            {
                featureLayer = (IFeatureLayer)GetLayer(aFeatureLayerName);
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("The layer " + aFeatureLayerName + " is not a feature layer");
                return false;
            }

            IActiveView activeView = GetActiveView();
            if (activeView == null || featureLayer == null || aWhereClause == null)
            {
                if (Messages)
                    MessageBox.Show("Please check input for this tool");
                return false;
            }

            // do this with Geoprocessor.
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            object sev = null;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(aFeatureLayerName);
            parameters.Add(aSelectionType);
            parameters.Add(aWhereClause);

            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("SelectLayerByAttribute_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);

                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                if (Messages)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(gp.GetMessages(ref sev));
                }
                gp = null;
                return false;
            }
            
        }

        public bool SelectLayerByLocation(string aTargetLayer, string aSearchLayer, string anOverlapType = "INTERSECT", string aSearchDistance = "", string aSelectionType = "NEW_SELECTION", bool Messages = false)
        {
            // Implementation of python SelectLayerByLocation_management.
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            object sev = null;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(aTargetLayer);
            parameters.Add(anOverlapType);
            parameters.Add(aSearchLayer);
            parameters.Add(aSearchDistance);
            parameters.Add(aSelectionType);

            // Clear selection. 
            if (aSelectionType == "NEW_SELECTION")
            {
                IFeatureSelection pFSel = (IFeatureSelection)GetLayer(aTargetLayer);
                pFSel.Clear();
            }

            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("SelectLayerByLocation_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);

                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                if (Messages)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(gp.GetMessages(ref sev));
                }
                gp = null;
                return false;
            }
        }

        public int CountSelectedLayerFeatures(string aFeatureLayerName, bool Messages = false)
        {
            // Check input.
            if (aFeatureLayerName == null)
            {
                if (Messages) MessageBox.Show("Please pass valid input string", "Feature Layer Has Selection");
                return -1;
            }

            if (!LayerExists(aFeatureLayerName))
            {
                if (Messages) MessageBox.Show("Feature layer " + aFeatureLayerName + " does not exist in this map");
                return -1;
            }

            IFeatureLayer pFL = null;
            try
            {
                pFL = (IFeatureLayer)GetLayer(aFeatureLayerName);
            }
            catch
            {
                if (Messages)
                    MessageBox.Show(aFeatureLayerName + " is not a feature layer");
                return -1;
            }

            IFeatureSelection pFSel = (IFeatureSelection)pFL;
            if (pFSel.SelectionSet.Count > 0) return pFSel.SelectionSet.Count;
            return 0;
        }

        public bool ClearSelection(string aFeatureLayerName, bool Messages = false)
        {
            if (!LayerExists(aFeatureLayerName))
                return false;

            IActiveView activeView = GetActiveView();
            IFeatureLayer featureLayer = null;
            try
            {
                featureLayer = (IFeatureLayer)GetLayer(aFeatureLayerName);
            }
            catch
            {
                if (Messages)
                    MessageBox.Show("The layer " + aFeatureLayerName + " is not a feature layer");
                return false;
            }

            if (activeView == null || featureLayer == null )
            {
                if (Messages)
                    MessageBox.Show("Please check input for this tool");
                return false;
            }

            // Clear selected features.
            IFeatureSelection pFSel = (IFeatureSelection)featureLayer;
            pFSel.Clear();
            return true;
        }

        public string GetOutputFileName(string aFileType, string anInitialDirectory = @"C:\")
        {
            // This would be done better with a custom type but this will do for the momment.
            IGxDialog myDialog = new GxDialogClass();
            myDialog.set_StartingLocation(anInitialDirectory);
            IGxObjectFilter myFilter;


            switch (aFileType)
            {
                case "Geodatabase FC":
                    myFilter = new GxFilterFGDBFeatureClasses();
                    break;
                case "Geodatabase Table":
                    myFilter = new GxFilterFGDBTables();
                    break;
                case "Shapefile":
                    myFilter = new GxFilterShapefiles();
                    break;
                case "DBASE file":
                    myFilter = new GxFilterdBASEFiles();
                    break;
                case "Text file":
                    myFilter = new GxFilterTextFiles();
                    break;
                default:
                    myFilter = new GxFilterDatasets();
                    break;
            }

            myDialog.ObjectFilter = myFilter;
            myDialog.Title = "Save Output As...";
            myDialog.ButtonCaption = "OK";

            string strOutFile = "None";
            if (myDialog.DoModalSave(thisApplication.hWnd))
            {
                strOutFile = myDialog.FinalLocation.FullName + @"\" + myDialog.Name;
                
            }
            myDialog = null;
            return strOutFile; // "None" if user pressed exit
            
        }

        #region CopyFeatures
        public bool CopyFeatures(string InFeatureClassOrLayer, string OutFeatureClass, string SortOrder, bool Messages = false)
        {
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = true;
            object sev = null;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(InFeatureClassOrLayer);
            parameters.Add(OutFeatureClass);

            // Execute the tool.
            try
            {
                if (SortOrder == "")
                {
                    myresult = (IGeoProcessorResult)gp.Execute("CopyFeatures_management", parameters, null);
                }
                else
                {
                    parameters.Add(SortOrder);
                    myresult = (IGeoProcessorResult)gp.Execute("Sort_management", parameters, null);
                }
                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);

                gp = null;
                myresult = null;
                parameters = null;
                return true;
            }
            catch (Exception ex)
            {
                if (Messages)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(gp.GetMessages(ref sev));
                }
                gp = null;
                return false;
            }
        }

        public bool CopyFeatures(string InWorkspace, string InDatasetName, string OutFeatureClass, string SortOrder, bool Messages = false)
        {
            string inFeatureClass = InWorkspace + @"\" + InDatasetName;
            return CopyFeatures(inFeatureClass, OutFeatureClass, SortOrder, Messages);
        }

        //public bool CopyFeatures(string InFeatureClass, string OutWorkspace, string OutDatasetName, string SortOrder, bool Messages = false)
        //{
        //    string outFeatureClass = OutWorkspace + @"\" + OutDatasetName;
        //    return CopyFeatures(InFeatureClass, outFeatureClass, SortOrder, Messages);
        //}

        public bool CopyFeatures(string InWorkspace, string InDatasetName, string OutWorkspace, string OutDatasetName, string SortOrder, bool Messages = false)
        {
            string inFeatureClass = InWorkspace + @"\" + InDatasetName;
            string outFeatureClass = OutWorkspace + @"\" + OutDatasetName;
            return CopyFeatures(inFeatureClass, outFeatureClass, SortOrder, Messages);
        }
        #endregion

        public bool CopyTable(string InTable, string OutTable, string SortOrder, bool Messages = false)
        {
            // This works absolutely fine for dbf and geodatabase but does not export to CSV.

            // Note the csv export already removes ghe geometry field; in this case it is not necessary to check again.

            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = true;
            object sev = null;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(InTable);
            parameters.Add(OutTable);

            // Execute the tool.
            try
            {
                if (SortOrder == "")
                {
                    myresult = (IGeoProcessorResult)gp.Execute("CopyRows_management", parameters, null);
                }
                else
                {
                    parameters.Add(SortOrder);
                    myresult = (IGeoProcessorResult)gp.Execute("Sort_management", parameters, null);
                }
                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);

                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                if (Messages)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(gp.GetMessages(ref sev));
                }
                gp = null;
                return false;
            }
        }

        public bool AddSpatialIndex(string InFeatureClassOrLayer, bool Messages = false)
        {
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = true;
            object sev = null;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(InFeatureClassOrLayer);

            // Execute the tool.
            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("AddSpatialIndex_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);

                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                if (Messages)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(gp.GetMessages(ref sev));
                }
                gp = null;
                return false;
            }
        }

        public bool AlterFieldAliasName(string aDatasetName, string aFieldName, string theAliasName, bool Messages = false)
        {
            // This script changes the field alias of a the named field in the layer.
            IObjectClass myObject = (IObjectClass)GetFeatureClass(aDatasetName);
            IClassSchemaEdit myEdit = (IClassSchemaEdit)myObject;
            try
            {
                myEdit.AlterFieldAliasName(aFieldName, theAliasName);
                myObject = null;
                myEdit = null;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                myObject = null;
                myEdit = null;
                return false;
            }
        }

        public IField getFCField(string InputDirectory, string FeatureclassName, string FieldName, bool Messages = false)
        {
            IFeatureClass featureClass = GetFeatureClass(InputDirectory, FeatureclassName);
            // Find the index of the requested field.
            int fieldIndex = featureClass.FindField(FieldName);

            // Get the field from the feature class's fields collection.
            if (fieldIndex > -1)
            {
                IFields fields = featureClass.Fields;
                IField field = fields.get_Field(fieldIndex);
                return field;
            }
            else
            {
                if (Messages)
                {
                    MessageBox.Show("The field " + FieldName + " was not found in the featureclass " + FeatureclassName);
                }
                return null;
            }
        }

        public IField getTableField(string TableName, string FieldName, bool Messages = false)
        {
            ITable theTable = GetTable(myFileFuncs.GetDirectoryName(TableName), myFileFuncs.GetFileName(TableName), Messages);
            int fieldIndex = theTable.FindField(FieldName);

            // Get the field from the feature class's fields collection.
            if (fieldIndex > -1)
            {
                IFields fields = theTable.Fields;
                IField field = fields.get_Field(fieldIndex);
                return field;
            }
            else
            {
                if (Messages)
                {
                    MessageBox.Show("The field " + FieldName + " was not found in the table " + myFileFuncs.GetFileName(TableName));
                }
                return null;
            }
        }

        public bool AppendTable(string InTable, string TargetTable, bool Messages = false)
        {
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = true;
            object sev = null;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();


            // Populate the variant array with parameter values.
            parameters.Add(InTable);
            parameters.Add(TargetTable);

            // Execute the tool. Note this only works with geodatabase tables.
            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("Append_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);

                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                if (Messages)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(gp.GetMessages(ref sev));
                }
                gp = null;
                return false;
            }
        }

        public bool DeleteTable(string InTable, bool Messages = false)
        {
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = true;
            object sev = null;

            IGeoProcessorResult myresult = new GeoProcessorResultClass();

            // Create a variant array to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(InTable);

            // Execute the tool. Note this only works with geodatabase tables.
            try
            {
                myresult = (IGeoProcessorResult)gp.Execute("Delete_management", parameters, null);

                // Wait until the execution completes.
                while (myresult.Status == esriJobStatus.esriJobExecuting)
                    Thread.Sleep(1000);

                gp = null;
                return true;
            }
            catch (Exception ex)
            {
                if (Messages)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(gp.GetMessages(ref sev));
                }
                gp = null;
                return false;
            }
        }

        public bool CopyToCSV(string InTable, string OutTable, List<string> aFieldList, bool Spatial, bool Append, bool Messages = false)
        {
            // This sub copies the input table to CSV.
            string aFilePath = myFileFuncs.GetDirectoryName(InTable);
            string aTabName = myFileFuncs.GetFileName(InTable);

            ICursor myCurs = null;
            IFields fldsFields = null;
            if (Spatial)
            {

                IFeatureClass myFC = GetFeatureClass(aFilePath, aTabName, true);
                myCurs = (ICursor)myFC.Search(null, false);
                fldsFields = myFC.Fields;
            }
            else
            {
                ITable myTable = GetTable(aFilePath, aTabName, true);
                myCurs = myTable.Search(null, false);
                fldsFields = myTable.Fields;
            }

            if (myCurs == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot open table " + InTable);
                }
                return false;
            }

            // Open output file.
            StreamWriter theOutput = new StreamWriter(OutTable, Append);

            string strHeader = "";
            int intFieldCount = fldsFields.FieldCount;
            List<Boolean> blIgnore = new List<Boolean>();

            // iterate through the fields in the collection to create header and determine which to ignore.
            int intFieldTot = 0;
            for (int i = 0; i < intFieldCount; i++)
            {
                // Get the field at the given index.
                IField pField = fldsFields.get_Field(i);
                string strField = pField.Name;

                // Does it exist in the 'keep' list or is it required?
                if ((aFieldList.IndexOf(strField) == -1) ||
                    (strField.ToUpper() == "SP_GEOMETRY") || (strField.ToLower() == "shape"))
                    blIgnore.Add(true);
                else
                {
                    blIgnore.Add(false);
                    strHeader = strHeader + strField + ",";
                    intFieldTot = intFieldTot + 1;
                }
            }

            if (!Append)
            {
                // Write the header.
                strHeader = strHeader.Substring(0, strHeader.Length - 1);
                theOutput.WriteLine(strHeader);
            }

            // Now write the file.
            IRow aRow = myCurs.NextRow();
            //MessageBox.Show("Writing ...");
            while (aRow != null)
            {
                string strRow = "";
                int intField = 0;
                for (int i = 0; i < intFieldCount; i++)
                {
                    if (blIgnore[i] == false)
                    {
                        var theValue = aRow.get_Value(i);
                        // Wrap value in quotes if it is a string that contains a comma.
                        if ((theValue is string) &&
                           (theValue.ToString().Contains(","))) theValue = "\"" + theValue.ToString() + "\"";
                        strRow = strRow + theValue.ToString();

                        intField = intField + 1;
                        if (intField < intFieldTot) strRow = strRow + ",";
                    }
                }
                theOutput.WriteLine(strRow);
                aRow = myCurs.NextRow();
            }

            theOutput.Close();
            theOutput.Dispose();
            myCurs = null;
            aRow = null;
            return true;
        }

        public bool WriteEmptyCSV(string OutTable, string theHeader)
        {
            // Open output file.
            StreamWriter theOutput = new StreamWriter(OutTable, false);
            theOutput.Write(theHeader);
            theOutput.Close();
            theOutput.Dispose();
            return true;
        }

        public bool CopyToTabDelimitedFile(string InTable, string OutTable, List<string> aFieldList, bool Spatial, bool Append, bool Messages = false)
        {
            // This sub copies the input table to CSV with no headers.
            string aFilePath = myFileFuncs.GetDirectoryName(InTable);
            string aTabName = myFileFuncs.GetFileName(InTable);

            ICursor myCurs = null;
            IFields fldsFields = null;
            if (Spatial)
            {

                IFeatureClass myFC = GetFeatureClass(aFilePath, aTabName, true);
                myCurs = (ICursor)myFC.Search(null, false);
                fldsFields = myFC.Fields;
            }
            else
            {
                ITable myTable = GetTable(aFilePath, aTabName, true);
                myCurs = myTable.Search(null, false);
                fldsFields = myTable.Fields;
            }

            if (myCurs == null)
            {
                if (Messages)
                {
                    MessageBox.Show("Cannot open table " + InTable);
                }
                return false;
            }

            // Open output file.
            StreamWriter theOutput = new StreamWriter(OutTable, Append);

            int intFieldCount = fldsFields.FieldCount;
            List<Boolean> blIgnore = new List<Boolean>();

            // iterate through the fields in the collection to determine which to ignore.
            int intFieldTot = 0;
            for (int i = 0; i < intFieldCount; i++)
            {
                // Get the field at the given index.
                IField pField = fldsFields.get_Field(i);
                string strField = pField.Name;

                // Does it exist in the 'keep' list or is it required?
                if ((aFieldList.IndexOf(strField) == -1) ||
                    (strField.ToUpper() == "SP_GEOMETRY") || (strField.ToLower() == "shape"))
                    blIgnore.Add(true);
                else
                {
                    blIgnore.Add(false);
                    intFieldTot = intFieldTot + 1;
                }
            }

            // Now write the file.
            IRow aRow = myCurs.NextRow();
            while (aRow != null)
            {
                string strRow = "";
                int intField = 0;
                for (int i = 0; i < intFieldCount; i++)
                {
                    if (blIgnore[i] == false)
                    {
                        var theValue = aRow.get_Value(i);
                        //// Wrap value in quotes if it is a string that contains a comma.
                        //if ((theValue is string) && (theValue.ToString().Contains(",")))
                        //    theValue = "\"" + theValue.ToString() + "\"";
                        // Wrap value in quotes.
                        strRow = strRow + "\"" + theValue.ToString() + "\"";

                        intField = intField + 1;
                        if (intField < intFieldTot) strRow = strRow + ",";
                    }
                }
                theOutput.WriteLine(strRow);
                aRow = myCurs.NextRow();
            }

            theOutput.Close();
            theOutput.Dispose();
            myCurs = null;
            aRow = null;
            return true;
        }

        public bool WriteEmptyTabDelimitedFile(string OutTable, string theHeader)
        {
            // Open output file.
            StreamWriter theOutput = new StreamWriter(OutTable, false);
            theHeader.Replace(",", "\t");
            theOutput.Write(theHeader);
            theOutput.Close();
            theOutput.Dispose();
            return true;
        }

        public void ToggleTOC()
        {
            IApplication m_app = thisApplication;

            IDockableWindowManager pDocWinMgr = m_app as IDockableWindowManager;
            UID uid = new UIDClass();
            uid.Value = "{368131A0-F15F-11D3-A67E-0008C7DF97B9}";
            IDockableWindow pTOC = pDocWinMgr.GetDockableWindow(uid);
            if (pTOC.IsVisible())
                pTOC.Show(false);
            else pTOC.Show(true);
            IMxApplication2 thisApp = (IMxApplication2)thisApplication;
            thisApp.Display.Invalidate(null, true, -1);
            IActiveView activeView = GetActiveView();
            activeView.Refresh();
            thisApplication.RefreshWindow();
        }

        public void ToggleDrawing()
        {
            IMxApplication2 thisApp = (IMxApplication2)thisApplication;
            thisApp.PauseDrawing = !thisApp.PauseDrawing;
            thisApp.Display.Invalidate(null, true, -1);
            IActiveView activeView = GetActiveView();
            activeView.Refresh();
            thisApplication.RefreshWindow();
        }

        public void ZoomToLayer(string aLayerName, bool Messages = false)
        {
            if (!LayerExists(aLayerName))
            {
                if (Messages)
                    MessageBox.Show("The layer " + aLayerName + " does not exist in the map");
                return;
            }
            IActiveView activeView = GetActiveView();
            ILayer pLayer = GetLayer(aLayerName);
            IEnvelope pEnv = pLayer.AreaOfInterest;
            pEnv.Expand(1.05, 1.05, true);
            activeView.Extent = pEnv;
            activeView.Refresh();
        }

        public void ZoomToFullExtent()
        {
            IActiveView activeView = GetActiveView();
            activeView.Extent = activeView.FullExtent;
            activeView.Refresh();
        }

        public IActiveView GetActiveView()
        {
            IMxDocument mxDoc = GetIMXDocument();
            return mxDoc.ActiveView;
        }
    }
}
