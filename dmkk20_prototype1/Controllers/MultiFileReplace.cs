using dmkk20_prototype1.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DirectoryModel = dmkk20_prototype1.Models.DirectoryModel;

// For protection removal

namespace dmkk20_prototype1.Controllers
{
    public class MultiFileReplace
    {
        private List<ChangeModel> Changes { get; set; }
        private List<ChangeModel> HeaderChanges { get; set; }
        private List<ChangeModel> FooterChanges { get; set; }
        private string CurrentDate { get; set; }
        private List<string> DirectoryList { get; set; }
        private List<string> FileList { get; set; }
        private LogFile Logger { get; set; }
        private string TemppFileName { get; set; }
        private string TempFileNameWithExt { get; set; }
        private Dictionary<string, string> Metadata { get; set; }
        public List<string> GetFileNames() { return FileList; }
        public List<string> GetDirectoryNames() { return DirectoryList; }
        public List<bool> HasChangedAnything { get; set; }

        public List<DirectoryModel> SelectedDocs;

        public MultiFileReplace(List<ChangeModel> changes, Dictionary<string, string> metadata, List<DirectoryModel> includedDocs, List<ChangeModel> headerChanges = null, List<ChangeModel> footerChanges = null)
        {
            Changes = changes;
            Metadata = metadata;
            HeaderChanges = headerChanges;
            FooterChanges = footerChanges;
            CurrentDate = DateTime.Now.ToString("dd-MM-yyyy");
            DirectoryList = new List<string>();
            FileList = new List<string>();
            HasChangedAnything = new List<bool>();

            SelectedDocs = includedDocs;
        }

        /*
         * Main function
         * Goes through searched directory
         * replaces the content through "ReadDcox"
         * unless if the file is logfile or deselected
         */
        public void GetDocx(string sDir)
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
                        if (TemppFileName != "logfile" && TemppFileName.Substring(0, 2) != "~$")
                        {
                            ReadDocx(file);
                        }
                    }
                    SearchSubFolders(sDir);
                }

                else
                {
                    // For each directory in root
                    foreach (string files in Directory.GetFiles(sDir))
                    {
                        TemppFileName = Path.GetFileNameWithoutExtension(files);
                        if (TemppFileName != "logfile" && TemppFileName.Substring(0, 2) != "~$")
                        {
                            ReadDocx(files);
                        }
                    }
                }
            }

            // Error catching
            catch (Exception ex)
            {
                throw new Exception("Failed to update the documents: " + ex.Message);
            }

        }


        /*
         * Called from main function for every existing file
         * which reads the given .docx files
         * and replaces the contents as required.
         */
        public void ReadDocx(string files)
        {
            // Separate Directory and file name (without extension name)
            string fileDir = Path.GetDirectoryName(files);
            string fileName = Path.GetFileNameWithoutExtension(files);
            string verify_full_fileDir = fileDir + "\\" + Path.GetFileName(files);

            List<bool> bodyMatch = new List<bool>();
            List<bool> headerMatch = new List<bool>();
            List<bool> footerMatch = new List<bool>();

            List<string> tempList = new List<string>();

            foreach (DirectoryModel row in SelectedDocs)
            {
                // If we should modify the document
                if (row.ApplyChanges)
                {
                    tempList.Add(row.DirectoryNames);
                }

            }

            // If current file is in list of to be modified documents
            if (tempList.Any(verify_full_fileDir.Contains))
            {
                byte[] byteArray = File.ReadAllBytes(files);
                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(byteArray, 0, (int)byteArray.Length);
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
                    {

                        DocumentProtection dp =
                                wordDoc.MainDocumentPart.DocumentSettingsPart.Settings.GetFirstChild<DocumentProtection>();
                        if (dp != null && dp.Enforcement == DocumentFormat.OpenXml.OnOffValue.FromBoolean(true))
                        {
                            dp.Remove(); // Doc is protected
                        }

                        foreach (var change in Changes)
                        {
                            if (HelperFunctions.AssertNotEmptyText(change.OldText, change.NewText))
                            {
                                bodyMatch.Add(ReplaceText(wordDoc, change.OldText, change.NewText));
                            }
                        }

                        foreach (var headerChange in HeaderChanges)
                        {
                            if (HelperFunctions.AssertNotEmptyText(headerChange.OldText, headerChange.NewText))
                            {
                                headerMatch.Add(ReplaceHeader(wordDoc, headerChange.OldText, headerChange.NewText));
                            }
                        }


                        foreach (var footerChange in FooterChanges)
                        {
                            if (HelperFunctions.AssertNotEmptyText(footerChange.OldText, footerChange.NewText))
                            {
                                footerMatch.Add(ReplaceFooter(wordDoc, footerChange.OldText, footerChange.NewText));
                            }
                        }

                        if (Metadata.Count > 0)
                        {
                            UpdateMetadata(wordDoc);
                        }

                    }

                    bool hasBodyMatch = bodyMatch.Any(x => x);
                    bool hasHeaderMatch = headerMatch.Any(x => x);
                    bool hasFooterMatch = footerMatch.Any(x => x);

                    if (hasBodyMatch || hasHeaderMatch || hasFooterMatch)
                    {
                        // List of files that are processed.
                        FileList.Add(fileName);

                        // Check if that file has a date at the end
                        string regexTest = @"_[0-9]{2}-[0-9]{2}-[0-9]{4}"; //check for _DD-MM-YYYY
                                                                           // If the file name contains date, it will take it out and replace for a new one, if not does nothing
                        fileName = Regex.Replace(fileName, regexTest, "");
                        fileName += "_" + CurrentDate + ".docx"; // Add new date with file etensions

                        // New file directory and old one
                        string new_fileDir = fileDir + "\\" + fileName;
                        string full_fileDir = fileDir + "\\" + Path.GetFileName(files);
                        // Replace the old by the new
                        string newPath = files.Replace(full_fileDir, new_fileDir);
                        DirectoryList.Add(full_fileDir); // Append to directory list
                        File.WriteAllBytes(newPath, stream.ToArray());

                        Logger = new LogFile(full_fileDir, Changes, hasBodyMatch, hasHeaderMatch, hasFooterMatch, HeaderChanges, FooterChanges);
                        Logger.CreateLogFile();

                        HasChangedAnything.Add(true);
                    }

                    else
                    {
                        HasChangedAnything.Add(false);
                    }
                }
            }
        }

        /*
         * Goes through each directory in the root path
         * Return all the document files that are not logfile and temporary edit word doc (~$)
         */
        public void SearchSubFolders(string sDir)
        {
            // For each directory in root
            foreach (string dir1 in Directory.GetDirectories(sDir))
            {
                // For each files of that specific directory
                foreach (string files1 in Directory.GetFiles(dir1))
                {
                    TemppFileName = Path.GetFileNameWithoutExtension(files1);
                    if (TemppFileName != "logfile")
                    {
                        ReadDocx(files1);
                    }
                }
                SearchSubFolders(dir1);
            }
        }


        /*
         * Replaces body contents
         */
        public bool ReplaceText(WordprocessingDocument doc, string oldText, string newText)
        {
            try
            {
                List<bool> bodyhasChanged = new List<bool>();

                var document = doc.MainDocumentPart.Document;

                oldText = Regex.Replace(oldText, @"[^\w\s]", @"\$&");
                Regex regexText = new Regex("\\b" + oldText.Replace(" ", "(.*?)") + "\\b");

                foreach (var text in document.Descendants<Text>()) 
                {
                    if (regexText.IsMatch(text.Text))
                    {                        
                        text.Text = regexText.Replace(text.Text, newText);
                        bodyhasChanged.Add(true);
                    }
                    else
                    {
                        bodyhasChanged.Add(false);
                    }
                }

                document.Save();

                // No replacement keywords found, don't perform any processing
                return bodyhasChanged.Any(x => x); 
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to Replace Text: " + ex.Message);
            }
        }

        /*
         *  Replaces header content
         */
        public bool ReplaceHeader(WordprocessingDocument doc, string oldText, string newText)
        {
            try
            {
                List<bool> headerhasChanged = new List<bool>();

                oldText = Regex.Replace(oldText, @"[^\w\s]", @"\$&");
                Regex regexText = new Regex("\\b" + oldText.Replace(" ", "(.*?)") + "\\b");

                foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                {
                    foreach(var currentText in headerPart.RootElement.Descendants<Text>())
                    {
                        if (regexText.IsMatch(currentText.Text))
                        {
                            currentText.Text = regexText.Replace(currentText.Text, newText);
                            headerhasChanged.Add(true);
                        }
                        else
                        {
                            headerhasChanged.Add(false);
                        }
                    }
                }
                return headerhasChanged.Any(x => x);

                
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to Replace the Header: " + ex.Message);
            }
        }

        /*
         * Replace footer content
         */
        public bool ReplaceFooter(WordprocessingDocument doc, string oldText, string newText)
        {
            try
            {
                List<bool> footerhasChanged = new List<bool>();

                oldText = Regex.Replace(oldText, @"[^\w\s]", @"\$&");
                Regex regexText = new Regex("\\b" + oldText.Replace(" ", "(.*?)") + "\\b");

                foreach (var footerPart in doc.MainDocumentPart.FooterParts)
                {
                    foreach (var currentText in footerPart.RootElement.Descendants<Text>())
                    {
                        if (regexText.IsMatch(currentText.Text))
                        {
                            currentText.Text = regexText.Replace(currentText.Text, newText);
                            footerhasChanged.Add(true);
                        }
                        else
                        { 
                            footerhasChanged.Add(false);
                        }
                    }
                }
                return footerhasChanged.Any(x => x);
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to Replace the Footer: " + ex.Message);
            }
        }


        /*
         * Updates metadata
         */
        public void UpdateMetadata(WordprocessingDocument doc)
        {
            try
            {
                string metadataText1 = null;
                using (StreamReader sr = new StreamReader(doc.CoreFilePropertiesPart.GetStream()))
                {
                    metadataText1 = sr.ReadToEnd();
                }

                string metadataText2 = null;
                using (StreamReader sr = new StreamReader(doc.ExtendedFilePropertiesPart.GetStream()))
                {
                    metadataText2 = sr.ReadToEnd();
                }

                foreach (var metadata in Metadata.Where(x => !string.IsNullOrEmpty(x.Value)))
                {
                    if (metadataText1.Contains(metadata.Key + ">"))
                    {
                        var substring = metadataText1.Substring(metadataText1.IndexOf(metadata.Key + ">"));
                        var value = substring.Substring(0, substring.IndexOf("</"));
                        metadataText1 = metadataText1.Replace(value, metadata.Key + ">" + metadata.Value);
                    }
                    if (metadataText2.Contains(metadata.Key + ">"))
                    {
                        var substring = metadataText2.Substring(metadataText2.IndexOf(metadata.Key + ">"));
                        var value = substring.Substring(0, substring.IndexOf("</"));
                        metadataText2 = metadataText2.Replace(value, metadata.Key + ">" + metadata.Value);
                    }
                }


                using (StreamWriter sw = new StreamWriter(doc.CoreFilePropertiesPart.GetStream(FileMode.Create)))
                {
                    sw.Write(metadataText1);
                }

                using (StreamWriter sw = new StreamWriter(doc.ExtendedFilePropertiesPart.GetStream(FileMode.Create)))
                {
                    sw.Write(metadataText2);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to Replace Text: " + ex.Message);
            }
        }

        /*
         * Indicates whether change has been made or not
         */
        public bool changesMade()
        {
            return HasChangedAnything.Any(x => x);
        }
    }
}
