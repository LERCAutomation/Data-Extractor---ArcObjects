﻿// DataSelector is an ArcGIS add-in used to extract biodiversity
// information from SQL Server based on any selection criteria.
//
// Copyright © 2016 Sussex Biodiversity Record Centre
//
// This file is part of DataSelector.
//
// DataSelector is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataSelector is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with DataSelector.  If not, see <http://www.gnu.org/licenses/>.


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace HLFileFunctions
{
    class FileFunctions
    {
        public bool DirExists(string aFilePath)
        {
            // Check input first.
            if (aFilePath == null) return false;
            DirectoryInfo myDir = new DirectoryInfo(aFilePath);
            if (myDir.Exists == false) return false;
            return true;
        }

        public string GetDirectoryName(string aFullPath)
        {
            // Check input.
            if (aFullPath == null) return null;

            // split at the last \
            int LastIndex = aFullPath.LastIndexOf(@"\");
            string aPath = aFullPath.Substring(0, LastIndex);
            return aPath;
        }

        public bool IsDirectory(string aPath)
        {
            FileAttributes attrThisPath = File.GetAttributes(aPath);
            if ((attrThisPath & FileAttributes.Directory) == FileAttributes.Directory)
                return true;
            else
                return false;
        }

        public List<string> GetSubdirectories(string aFullPath)
        {
            
            if (!IsDirectory(aFullPath))
                aFullPath = GetDirectoryName(aFullPath);

            List<string> liSubDirs = aFullPath.Split('\\').ToList() ;//Directory.GetDirectories(aFullPath).ToList();

            // Cycle through 

            return liSubDirs;

        }

        public string GetRelativeSubdirectories(string aFullPath, string aBaseFolder)
        {
            // This function returns the subdirectories in aFullPath relative to aBaseFolder.
            // Example: c:\aHester\Projects\ThisProject with aBaseFolder c:\aHester returns Projects\ThisProject
            // The same with aBaseFolder c:\aHester\Projects returns ThisProject

            // Firstly get all the subdirectories in the FullPath
            List<string> liFullDirs = GetSubdirectories(aFullPath);
            // Do the same for the BaseFolder.
            List<string> liBaseDirs = GetSubdirectories(aBaseFolder);

            // Do a few basic checks.
            if (liFullDirs.Count < liBaseDirs.Count)
                return null; // Cannot be subdirectories.
            int i = 0;
            foreach (string aBaseSubdir in liBaseDirs)
            {
                if (liFullDirs[i] != aBaseSubdir)
                    return null; // Is not a subdirectory as it doesn't follow the same tree.
                i++;
            }


            // What is the last folder name in the base?
            string strLastBase = liBaseDirs[liBaseDirs.Count - 1];
            // what is the index of this folder name in the full path?
            i = liFullDirs.IndexOf(strLastBase);
            // Now build the return string.
            string strOutput = "";
            if (i < liFullDirs.Count - 1) // If there is a subdirectory at all.
            {
                for (int a = i+1; a < liFullDirs.Count(); a++)
                {
                    if (strOutput == "")
                        strOutput = liFullDirs[a];
                    else
                        strOutput = strOutput + @"\" + liFullDirs[a];
                }
            }
            return strOutput;
        }

        #region FileExists
        public bool FileExists(string aFilePath, string aFileName)
        {
            if (DirExists(aFilePath))
            {
                string strFileName = aFilePath;
                string aTest = aFilePath.Substring(aFilePath.Length - 1, 1);
                if (aTest != @"\")
                {
                    strFileName = strFileName + @"\" + aFileName;
                }
                else
                {
                    strFileName = strFileName + aFileName;
                }

                System.IO.FileInfo myFileInfo = new FileInfo(strFileName);

                if (myFileInfo.Exists) return true;
                else return false;
            }
            return false;
        }
        public bool FileExists(string aFullPath)
        {
            System.IO.FileInfo myFileInfo = new FileInfo(aFullPath);
            if (myFileInfo.Exists) return true;
            return false;
        }

        #endregion

        public string GetFileName(string aFullPath)
        {
            // Check input.
            if (aFullPath == null) return null;

            // split at the last \
            int LastIndex = aFullPath.LastIndexOf(@"\");
            string aFile = aFullPath.Substring(LastIndex + 1, aFullPath.Length - (LastIndex + 1));
            return aFile;
        }

        public List<string> GetAllFilesInfolder(string aFullPath)
        {
            DirectoryInfo myDir = new DirectoryInfo(aFullPath);
            FileInfo[] allFilesInDir = myDir.GetFiles();
            List<string> liAllFiles = new List<string>();
            foreach (FileInfo aThing in allFilesInDir)
            {
                string strName = aThing.Name.ToLower();
                if(!strName.Contains(".lock") ) // Ignore ESRI lock files
                {
                   liAllFiles.Add(aThing.FullName);
                }
            }

            // Any subdirectories?
            DirectoryInfo[] allSubDirs = myDir.GetDirectories();
            foreach (DirectoryInfo aDir in allSubDirs)
            {
                allFilesInDir = aDir.GetFiles();
                foreach (FileInfo aThing in allFilesInDir)
                {
                    string strName = aThing.Name.ToLower();
                    if (!strName.Contains(".lock")) // Ignore ESRI lock files
                    {
                        liAllFiles.Add(aThing.FullName);
                    }
                }
            }
            return liAllFiles;
        }
    

        public string ReturnWithoutExtension(string aFileName)
        {
            // check input
            if (aFileName == null) return null;
            int aLen = aFileName.Length;
            // check if it has an extension at all
            string aTest = aFileName.Substring(aLen - 4, 1);
            if (aTest != ".") return aFileName;

            return aFileName.Substring(0, aLen - 4);
        }

        public bool DeleteFile(string aFullPath)
        {
            if (FileExists(aFullPath))
            {
                try
                {
                    File.Delete(aFullPath);
                    return true;
                }
                catch
                {
                    return false;
                }
            }
            else
                return true;

        }

        public bool CreateLogFile(string aTextFile)
        {
            StreamWriter myWriter = new StreamWriter(aTextFile, false);

            myWriter.WriteLine("Log file for Data Selector, started on " + DateTime.Now.ToString());
            myWriter.Close();
            myWriter.Dispose();
            return true;
        }

        public bool WriteLine(string aTextFile, string aWriteLine)
        {
            StreamWriter myWriter = new StreamWriter(aTextFile, true);
            aWriteLine = DateTime.Now.ToString() + " : " + aWriteLine;
            myWriter.WriteLine(aWriteLine);
            myWriter.Close();
            myWriter.Dispose();
            return true;
        }
    }
}
