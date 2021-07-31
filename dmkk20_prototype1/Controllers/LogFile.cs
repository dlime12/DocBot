using dmkk20_prototype1.Models;
using System;
using System.Collections.Generic;
using System.IO;
using Xceed.Words.NET;

namespace dmkk20_prototype1.Controllers
{
    public class LogFile
    {
        private const string fileName = "\\logfile.docx";
        private string Path { get; set; }
        private List<ChangeModel> BodyChanges {get; set;}
        private string FileName { get; set; }
        private List<ChangeModel> HeaderChanges { get; set; }
        private List<ChangeModel> FooterChanges { get; set; }
        private string NewLog { get; set; }
        private bool HasChangedBody { get; set; }
        private bool HasChangedHeader { get; set; }
        private bool HasChangedfooter { get; set; }

        public LogFile(string path, List<ChangeModel> changes, bool hasChangedBody, bool hasChangedHeader, bool hasChangedFooter, List<ChangeModel> headerChanges = null, List<ChangeModel> footerChanges = null)
        {
            BodyChanges = changes;
            HeaderChanges = headerChanges;
            FooterChanges = footerChanges;
            FileName = System.IO.Path.GetFileName(path);
            Path = System.IO.Path.GetDirectoryName(path) + fileName;
            HasChangedBody = hasChangedBody;
            HasChangedHeader = hasChangedHeader;
            HasChangedfooter = hasChangedFooter;

        }

        public void CreateLogFile()
        {
            if (File.Exists(Path))
            {
                DocX ExistingLogFile = DocX.Load(Path);

                CreateLog();

                NewLog = Environment.NewLine + NewLog;
                ExistingLogFile.InsertParagraph(NewLog);
                ExistingLogFile.Save();
            }
            else
            {
                // Create the log file document
                var logFile = DocX.Create(Path);

                CreateLog();

                // Add the log to the document and save
                logFile.InsertParagraph(NewLog);
                logFile.Save();
            }
        }

        public void CreateLog()
        {
            //body
            if (HasChangedBody)
            {
                foreach (var change in BodyChanges)
                {
                    if (HelperFunctions.AssertNotEmptyText(change.OldText, change.NewText))
                        GenerateLogText("Body", change.OldText, change.NewText);
                }
            }
            

            //header
            if (HasChangedHeader)
            {
                foreach (var change in HeaderChanges)
                {
                    if (ValidateInputs(change.OldText, change.NewText))
                        GenerateLogText("Header", change.OldText, change.NewText);
                }
            }
            


            //footer
            if (HasChangedfooter)
            {
                foreach (var change in FooterChanges)
                {
                    if (ValidateInputs(change.OldText, change.NewText))
                        GenerateLogText("Footer", change.OldText, change.NewText);
                }
            }       
        }

        private string GenerateLogText(string @field, string old, string @new)
        {
            return NewLog += $"\n{DateTime.Now}: replaced {@field} text '{old}' to '{@new}' in the file '{FileName}'";
        }

        private bool AssertNotDefaultModelText(string oldText, string newText)
        {
            return !(oldText == "Old text" || newText == "New text");
        }

        private bool ValidateInputs(string oldText, string newText)
        {
            return AssertNotDefaultModelText(oldText, newText) && HelperFunctions.AssertNotEmptyText(oldText, newText);
        }

    }
}
