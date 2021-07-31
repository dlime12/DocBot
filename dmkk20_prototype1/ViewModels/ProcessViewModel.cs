using dmkk20_prototype1.Controllers;
using dmkk20_prototype1.Models;
using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Navigation;
using static dmkk20_prototype1.Views.MainWindow;

namespace dmkk20_prototype1.ViewModels
{
    public class ProcessViewModel
    {
        public string Path { get; set; }
        public List<ChangeModel> HeaderChanges { get; set; }
        public List<ChangeModel> FooterChanges {get; set; }
        public List<ChangeModel> Changes { get; set; }
        public List<string> ResultFiles { get; set; }
        public List<string> ResultDirectories { get; set; }
        public Dictionary<string, string> Metadata { get; set; }

        public List<DirectoryModel> documentsToBeChanged;

        public ProcessViewModel()
        {
            Changes = new List<ChangeModel>
            {
                new ChangeModel()
            };
            HeaderChanges = new List<ChangeModel>
            {
                new ChangeModel()
            };
            FooterChanges = new List<ChangeModel>
            {
                new ChangeModel()
            };
            Metadata = new Dictionary<string, string>();
        }


        // Search button is pressed => Return all the .docx files in the given directory for result panel
        public bool ReturnFiles()
        {
            bool directoryValid = false;

            // Validate input
            if (string.IsNullOrEmpty(Path))
            {
                MessageBox.Show("Path cannot be empty", "Empty directory path");
                return directoryValid;
            }
           
            // Get files
            try
            {
                // Search the directory for document files
                MultiFileSearch multiSearch = new MultiFileSearch();
                multiSearch.InitialiseFiles(Path);
                ResultFiles = multiSearch.GetFileNames();
                ResultDirectories = multiSearch.GetDirectoryNames();
            }
            catch (Exception criticalError)
            {
                MessageBox.Show("Could not perform file search.\n\n Error Message: " + criticalError, "Critical error");
                return directoryValid;
            }

            // Check if search results returned any .docx files
            if (ResultFiles.Any())
            {
                directoryValid = true;
            }
            else
            { 
                MessageBox.Show("There aren't any documents in this folder. \nDid you input the correct directory?", "No files found");
            }
            return directoryValid;
        }


        // Replace button is pressed => Process all the .docx changes to the selected documents
        public bool Process()
        {
            bool updateStatus = false;

            // Validate Path
            if (string.IsNullOrEmpty(Path))
            {
                MessageBox.Show("Path cannot be empty", "Empty directory path");
                return updateStatus;
            }

            try
            {
                // Multi file replace 
                MultiFileReplace multiReplace = new MultiFileReplace(Changes, Metadata, documentsToBeChanged, HeaderChanges, FooterChanges);
                multiReplace.GetDocx(Path);                              
                
                if (multiReplace.changesMade())
                {
                    MessageBox.Show("Document/s Successfully Modified", "Success");
                }                
                else
                {
                    MessageBox.Show("Search terms were not found in the specified documents.", "No changes made");
                }
                return updateStatus = true;
            }
            catch (Exception criticalError)
            {
                MessageBox.Show("Could not update documents.\n\n Error Message: " + criticalError, "Critical error");
                return updateStatus;
            }
        }
    }
}
