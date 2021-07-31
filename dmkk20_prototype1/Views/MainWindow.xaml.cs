using dmkk20_prototype1.Models;
using dmkk20_prototype1.ViewModels;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using WinForms = System.Windows.Forms;

namespace dmkk20_prototype1.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ProcessViewModel viewModel;

        public MainWindow()
        {
            viewModel = new ProcessViewModel();
            InitializeComponent();
        }


        // #################################
        // ##### Input label processing #####
        // #################################
        
        
        // Input Directory label updates
        private void Directory_TextChanged(object sender, TextChangedEventArgs e)
        {
            // If the directory input isn't empty
            if(sender != null)
            {
                // Retrieve the directory and set it to our base path to our ProcessViewModel object.
                var textBox = (TextBox)sender;
                viewModel.Path = textBox.Text;
            }
        }

        // update the header text that has to be changed
        private void OldHeader_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender != null)
            {
                var textBox = (TextBox)sender;
                viewModel.HeaderChanges[Convert.ToInt32(textBox.Tag) - 1].OldText = textBox.Text;
            }
        }

        // update the new header text
        private void NewHeader_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender != null)
            {
                var textBox = (TextBox)sender;
                viewModel.HeaderChanges[Convert.ToInt32(textBox.Tag) - 1].NewText = textBox.Text;
            }
        }

        // update the footer text that has to be changed
        private void OldFooter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender != null)
            {
                var textBox = (TextBox)sender;
                viewModel.FooterChanges[Convert.ToInt32(textBox.Tag) - 1].OldText = textBox.Text;
            }
        }

        // update the new footer text
        private void NewFooter_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender != null)
            {
                var textBox = (TextBox)sender;
                viewModel.FooterChanges[Convert.ToInt32(textBox.Tag) - 1].NewText = textBox.Text;
            }
        }

        // update the ReplaceFrom text
        private void Search_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender != null)
            {
                var textBox = (TextBox)sender;
                viewModel.Changes[Convert.ToInt32(textBox.Tag) - 1].OldText = textBox.Text;
            }
        }

        // update the ReplaceTo text
        private void Replace_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender != null)
            {
                var textBox = (TextBox)sender;
                viewModel.Changes[Convert.ToInt32(textBox.Tag) - 1].NewText = textBox.Text;
            }
        }

        // Update meta-data texts
        public void Metadata_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender != null)
            {
                var textBox = (TextBox)sender;
                if (viewModel.Metadata.ContainsKey(textBox.Name))
                {
                    viewModel.Metadata[textBox.Name] = textBox.Text;
                }
                else
                {
                    viewModel.Metadata.Add(textBox.Name, textBox.Text);
                }
            }
        }


        // #################################
        // ##### Buttons and functions #####
        // #################################


        // choose the directory path and update the directory text box
        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            using(var dialog = new WinForms.FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    DirectoryInput.Text = dialog.SelectedPath;
                }
            }
        }
        
        // Search first
        private void Search_Click(object sender, RoutedEventArgs e)
        {
            bool directoryValid = viewModel.ReturnFiles();

            if (directoryValid)
            {
                // Create a new list of directory information
                viewModel.documentsToBeChanged = new List<DirectoryModel>();
                // Get the list of file/s and directory/ies
                List<string> fileNameList = viewModel.ResultFiles;
                List<string> directoryNameList = viewModel.ResultDirectories;

                // Iterate through the list and append to the documents to be changed
                for (var index = 0; index < fileNameList.Count; index++)
                {
                    // Default as selected, filename and file directory.
                    viewModel.documentsToBeChanged.Add(new DirectoryModel()
                    {
                        ApplyChanges = true,
                        FileNames = fileNameList[index],
                        DirectoryNames = directoryNameList[index]
                    });
                }

                resultTable.ItemsSource = viewModel.documentsToBeChanged;
                resultsLabel.Content = "List panel last updated - " + DateTime.Now.ToString("MMMM dd") + " " + DateTime.Now.ToString("HH:mm:ss") + " " + DateTime.Now.ToString("tt");
                resultsLabel.Foreground = Brushes.Green;
            }
            else
            {
                resultsLabel.Content = "Please input a valid directory!";
                resultsLabel.Foreground = Brushes.Red;
            }
        }

        // search and replace the text by the click of the button
        // by the end creates a message box to display the result
        private void Replace_Click(object sender, RoutedEventArgs e)
        {
            // Search hasn't been clicked or no documents
            if (viewModel.documentsToBeChanged == null)
            {
                MessageBox.Show("There are no documents to modify.", "No documents found");
            }

            // Check if any selections were made
            else if (!viewModel.documentsToBeChanged.Any(x => x.ApplyChanges == true))
            {
                MessageBox.Show("No documents were selected.", "No document selection");
            }

            // Verify modification and then execute.
            else 
            {
                MessageBoxResult result = MessageBox.Show("Do you want to apply the changes?",
                                              "Confirmation",
                                              MessageBoxButton.YesNo,
                                              MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    // ProcessViewModel call
                    viewModel.Process();
                }
            }
        }

        // If a cell is clicked in the DirectoryNames column, open the corresponding word document.
        private void DataGridCell_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Get the specifics of the currently clicked cell
            var dataGridCellTarget = (DataGridCell)sender;
            string selectedCell = dataGridCellTarget.ToString();

            // Get the .docx path and then open the word document file
            string trimmedSelected = selectedCell.Substring(38);
            System.Diagnostics.Process.Start(trimmedSelected);
            MessageBox.Show("Document opening... Make sure to close any opened word documents before executing DocBot");
        }
        
        // add search and replace fields
        private void AddReplaceFields_Click(object sender, RoutedEventArgs e)
        {
            var button = e.Source as Button;
            var canvas = button.Parent as Canvas;
            var topLoc = Canvas.GetTop(button);
            canvas.Height += 50;
            Canvas.SetTop(button, topLoc + 50);

            // Create an instance of a new next box for old text
            var search = new TextBox();
            search.HorizontalAlignment = HorizontalAlignment.Left;
            search.TextWrapping = TextWrapping.Wrap;
            search.VerticalAlignment = VerticalAlignment.Top;
            search.Text = "Old text";
            search.Height = 30;
            search.Width = 158;
            Canvas.SetTop(search, topLoc);
            Canvas.SetLeft(search, 6.00);

            // Create an instance of a new next box for new text
            var replace = new TextBox();
            replace.HorizontalAlignment = HorizontalAlignment.Left;
            replace.TextWrapping = TextWrapping.Wrap;
            replace.VerticalAlignment = VerticalAlignment.Top;
            replace.Text = "New text";
            replace.Height = 30;
            replace.Width = 158;
            Canvas.SetTop(replace, topLoc);
            Canvas.SetLeft(replace, 183.00);

            switch (canvas.Tag.ToString())
            {
                case "Body":
                    search.Tag = (viewModel.Changes.Count + 1).ToString();
                    search.TextChanged += new TextChangedEventHandler(Search_TextChanged);
                    replace.Tag = (viewModel.Changes.Count + 1).ToString();
                    replace.TextChanged += new TextChangedEventHandler(Replace_TextChanged);
                    viewModel.Changes.Add(new ChangeModel());
                    break;
                case "Header":
                    search.Tag = (viewModel.HeaderChanges.Count + 1).ToString();
                    search.TextChanged += new TextChangedEventHandler(OldHeader_TextChanged);
                    replace.Tag = (viewModel.HeaderChanges.Count + 1).ToString();
                    replace.TextChanged += new TextChangedEventHandler(NewHeader_TextChanged);
                    viewModel.HeaderChanges.Add(new ChangeModel());
                    break;
                case "Footer":
                    search.Tag = (viewModel.FooterChanges.Count + 1).ToString();
                    search.TextChanged += new TextChangedEventHandler(OldFooter_TextChanged);
                    replace.Tag = (viewModel.FooterChanges.Count + 1).ToString();
                    replace.TextChanged += new TextChangedEventHandler(NewFooter_TextChanged);
                    viewModel.FooterChanges.Add(new ChangeModel());
                    break;
            }

            // Add items to the UI
            canvas.Children.Add(search);
            canvas.Children.Add(replace);
        }
    }
}
