using dmkk20_prototype1.Models;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DirectoryModel = dmkk20_prototype1.Models.DirectoryModel;
namespace dmkk20_prototype1.Controllers
{
    class MultiFileSearch
    {
        private List<string> DirectoryList { get; set; }
        private List<string> FileList { get; set; }
        private string TemppFileName { get; set; }
        private string TempFileNameWithExt { get; set; }
        public List<string> GetFileNames() { return FileList; }
        public List<string> GetDirectoryNames() { return DirectoryList; }

        public MultiFileSearch()
        {
            DirectoryList = new List<string>();
            FileList = new List<string>();
        }

        /*
         * Returns all the found .docx files within a directory
         */
        public void GetFileDirectories(string files)
        {
            // Separate Directory and file name (without extension name)
            string fileDir = Path.GetDirectoryName(files);
            string fileName = Path.GetFileNameWithoutExtension(files);
            FileList.Add(fileName); // Append to modified file list 
                                    // Check if that file has a date at the end

            // New file directory and old one
            string full_fileDir = fileDir + "\\" + Path.GetFileName(files);

            DirectoryList.Add(full_fileDir); // Append to directory list
        }

        /*
         * Goes through each directory in the root path
         * Return all the document files that are not logfile and temporary edit word doc (~$)
         */
        public void multiDir(string strDir)
        {
            // For each directory in root
            foreach (string currDirectory in Directory.GetDirectories(strDir))
            {
                // For each files of that specific directory
                foreach (string currFile in Directory.GetFiles(currDirectory))
                {
                    TemppFileName = Path.GetFileNameWithoutExtension(currFile);
                    TempFileNameWithExt = Path.GetFileName(currFile);
                    if (TemppFileName != "logfile" && TemppFileName.Substring(0,2) != "~$")
                    {
                        if (TempFileNameWithExt.Contains(".docx"))
                        {
                            GetFileDirectories(currFile);
                        }
                    }
                }
                multiDir(currDirectory);
            }
        }

        /*
         * Go through the root folder and find all .docx files
         */
        public void InitialiseFiles(string sDir)
        {
            // Try to read through every directory and files nested inside the root
            try
            {
                string[] directoriesExists = Directory.GetDirectories(sDir);
                if (directoriesExists.Length > 0)
                {
                    foreach (string file in Directory.GetFiles(sDir))
                    {
                        TemppFileName = Path.GetFileNameWithoutExtension(file);
                        TempFileNameWithExt = Path.GetFileName(file);
                        if (TemppFileName != "logfile" && TemppFileName.Substring(0, 2) != "~$")
                        {
                            if (TempFileNameWithExt.Contains(".docx"))
                            {
                                GetFileDirectories(file);
                            }
                        }
                    }
                    multiDir(sDir);
                }
                else
                {
                    // For each directory in root
                    foreach (string files in Directory.GetFiles(sDir))
                    {
                        TemppFileName = Path.GetFileNameWithoutExtension(files);
                        TempFileNameWithExt = Path.GetFileName(files);
                        if (TemppFileName != "logfile" && TemppFileName.Substring(0,2) != "~$")
                        {
                            if (TempFileNameWithExt.Contains(".docx"))
                            {
                                GetFileDirectories(files);
                            }
                        }
                    }
                }
            }

            // Error catching
            catch (Exception ex)
            {
                throw new Exception("Failed while searching for documents: " + ex.Message);
            }
        }

    }
}
