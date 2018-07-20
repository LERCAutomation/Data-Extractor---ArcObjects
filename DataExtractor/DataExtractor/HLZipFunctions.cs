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
//
// Code used is from http://www.csharpque.com/2013/05/zipfiles-dotnet35-CreateAssembly.html
// Spellchecked and extended by Hester Lyons.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Packaging;
using Shell32;
using HLFileFunctions;

using System.Windows.Forms;

// FACTORY CLASS
public abstract class ArchiveObj
{
    internal String strPath;
    internal List<String> lError;
    internal List<String> lFiles;
    internal List<String> lFolders;
    internal FileFunctions myFileFuncs = new FileFunctions();
    public String[] ErrorList
    {
        get
        {
            return lError.ToArray();
        }
    }
    public ArchiveObj()
    {
        lFiles = new List<string>();
        lFolders = new List<string>();
        lError = new List<string>();
    }
    public Boolean AddFile(String strFile, string strBaseFolder)
    {
        lFiles.Add(strFile);

        // Work out what its base folder should be. If none, add an empty string.
        string strSubDirs = myFileFuncs.GetRelativeSubdirectories(strFile, strBaseFolder);
        
        if (strSubDirs != "")
            strSubDirs = "/" + strSubDirs;
        lFolders.Add(strSubDirs);
        return true;
    }

    public Boolean AddAllFiles(String strFolder)
    {
        // Get all files in this folder
        List<string> liFiles = myFileFuncs.GetAllFilesInfolder(strFolder);
        foreach (string strFile in liFiles)
        {
            AddFile(strFile, strFolder);
        }
        return true;
    }

    public abstract int SaveArchive();
}
// CREATOR CLASS
public abstract class ArchiveCreator
{
    internal String strPath;
    public abstract ArchiveObj GetArchive();
}



namespace Archive
{
    class WinBaseZip : ArchiveObj 
    {
        private WinBaseZip(){}
        public WinBaseZip(String sPath)
        {
            strPath = sPath;
        }
        public override int SaveArchive()
        {
            int i = 0;
            foreach (String strFile in lFiles)
            {
                zipFile(strFile, lFolders[i]);
                i++;
            }
            if (lError.Count > 0)
            {
                if (lError.Count < lFiles.Count)
                    return 0;
                else
                    return -1;
            }
           return 1;
        }
        private bool zipFile(String strFile, string strSubDir) // Change this so that it accepts subdirectories.
        {
            PackagePart pkgPart = null;
            using (Package Zip = Package.Open(strPath, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite))
            {
                String strTemp = strFile.Replace(" ", "_"); // We amended to use subdirectories.
                String zipURI = String.Concat(strSubDir,"/", System.IO.Path.GetFileName(strTemp));
                Uri parturi = new Uri(zipURI, UriKind.Relative);
                try
                {
                    pkgPart = Zip.CreatePart(parturi, System.Net.Mime.MediaTypeNames.Application.Zip, CompressionOption.Normal);
                }
                catch(Exception ex)
                {
                    lError.Add(strFile + "; Error : " + ex.Message);
                    return false;
                }
                Byte[] bites = System.IO.File.ReadAllBytes(strFile);
                pkgPart.GetStream().Write(bites, 0, bites.Length);
                Zip.Close();
            }
            return true;
        }
    }

    public class WinBaseZipCreator : ArchiveCreator
    {
        private WinBaseZipCreator(){}
        public WinBaseZipCreator(String sPath)
        {
            strPath = sPath;
        }
        public override ArchiveObj GetArchive()
        {
            return new WinBaseZip(strPath);
        }
    }

    
}
namespace Archive
{
    class ZipShell : ArchiveObj
    {
        private ZipShell() { }
        public ZipShell(String sPath)
        {
            strPath = sPath;
        }
        public override int SaveArchive()
        {
            CreateZip(strPath);
            foreach (String strFile in lFiles)
                ZipCopyFile(strFile);
            return 1;
        }
        public bool CreateZip(String strZipFile)
        {
            try
            {
                System.Text.ASCIIEncoding Encoder = new System.Text.ASCIIEncoding();
                byte[] baHeader = System.Text.Encoding.ASCII.GetBytes(("PK" + (char)5 + (char)6).PadRight(22, (char)0));
                System.IO.FileStream fs = System.IO.File.Create(strZipFile);
                fs.Write(baHeader, 0, baHeader.Length);
                fs.Flush();
                fs.Close();
                fs = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating zip file. The system returned the following message: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return true;
        }
        private bool ZipCopyFile(String strFile)
        {
            Shell32.Shell Shell = new Shell32.Shell();
            int iCnt = Shell.NameSpace(strPath).Items().Count;
            Shell.NameSpace(strPath).CopyHere(strFile, 0); // Copy file in Zip
            if (Shell.NameSpace(strPath).Items().Count == (iCnt + 1))
            {
                System.Threading.Thread.Sleep(100);
            }
            return true;
        }
    }
    public class ZipShellCreator : ArchiveCreator
    {
        private ZipShellCreator() { }
        public ZipShellCreator(String sPath)
        {
            strPath = sPath;
        }
        public override ArchiveObj GetArchive()
        {
            return new ZipShell(strPath);
        }
    }
}