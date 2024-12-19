using Microsoft.VisualBasic.FileIO;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace CSVProcessorApp
{
    public partial class MainWindow : Window
    {
        private string? filePath; // Nullable to handle initialization

        public MainWindow()
        {
            InitializeComponent();
            OfficeOpenXml.ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set EPPlus license
        }

        private void OnAboutClick(object sender, RoutedEventArgs e)
        {
            var aboutWindow = new AboutWindow();
            aboutWindow.ShowDialog();
        }

        private void OnBrowseClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                Title = "Select a CSV file"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                filePath = openFileDialog.FileName;
                FilePathTextBox.Text = filePath;
            }
        }

        private void OnFileDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    filePath = files[0];
                    FilePathTextBox.Text = filePath;
                }
            }
        }

        private void OnProcessClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("Please select a valid CSV file.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // Parse CSV file
                var lines = ParseCsvFile(filePath);

                // Locate the header row and "Total" column
                int headerRowIndex = -1;
                int totalColumnIndex = -1;
                for (int i = 0; i < lines.Count; i++)
                {
                    totalColumnIndex = Array.IndexOf(lines[i], "Total");
                    if (totalColumnIndex != -1)
                    {
                        headerRowIndex = i;
                        break;
                    }
                }

                if (headerRowIndex == -1 || totalColumnIndex == -1)
                {
                    MessageBox.Show("The CSV file does not contain a 'Total' column.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Sum the values in the "Total" column
                double grandTotal = lines.Skip(headerRowIndex + 1)
                                         .Where(row => row.Length > totalColumnIndex)
                                         .Sum(row => double.TryParse(row[totalColumnIndex].Replace("$", ""), out var value) ? value : 0);


                // Show Gratuity Input Window
                var gratuityInputWindow = new GratuityInputWindow();
                if (gratuityInputWindow.ShowDialog() == true)
                {
                    double gratuityRate = gratuityInputWindow.GratuityPercentage / 100.0;
                    double gratuityAmount = grandTotal * gratuityRate;
                    double finalTotal = grandTotal + gratuityAmount;

                    // Create Excel file with original data
                    SaveFileDialog saveFileDialog = new SaveFileDialog
                    {
                        Filter = "Excel Macro-Enabled files (*.xlsx)|*.xlsx",
                        Title = "Save Processed File"
                    };

                    if (saveFileDialog.ShowDialog() == true)
                    {
                        var saveFilePath = saveFileDialog.FileName;
                        var fileInfo = new FileInfo(saveFilePath);

                        if (fileInfo.Exists)
                        {
                            fileInfo.Delete(); // Delete the existing file to allow overwriting
                        }

                        using (var package = new ExcelPackage(new FileInfo(saveFilePath)))
                        {
                            var worksheet = package.Workbook.Worksheets.Add("Processed Data");

                            // Write original data to the worksheet
                            for (int i = 0; i < lines.Count; i++)
                            {
                                for (int j = 0; j < lines[i].Length; j++)
                                {
                                    worksheet.Cells[i + 1, j + 1].Value = lines[i][j];
                                }
                            }

                            // Add "Grand Total" row, shifted to columns K and L
                            int totalRow = lines.Count + 1;
                            worksheet.Cells[totalRow, 11].Value = "Grand Total"; // Column K (11)
                            worksheet.Cells[totalRow, 11].Style.Font.Bold = true; // Make the label bold
                            worksheet.Cells[totalRow, 12].Value = grandTotal;   // Column L (12)
                            worksheet.Cells[totalRow, 12].Style.Font.Bold = true; // Make the value bold
                            worksheet.Cells[totalRow, 12].Style.Numberformat.Format = "$#,##0.00"; // Format with dollar sign

                            //Add provided gratuity calculation
                            worksheet.Cells[totalRow + 1, 11].Value = $"{gratuityInputWindow.GratuityPercentage}% Gratuity";
                            worksheet.Cells[totalRow + 1, 11].Style.Font.Bold = true;
                            worksheet.Cells[totalRow + 1, 12].Value = gratuityAmount;
                            worksheet.Cells[totalRow + 1, 12].Style.Font.Bold = true;
                            worksheet.Cells[totalRow + 1, 12].Style.Numberformat.Format = "$#,##0.00";

                            //Final total including gratuity
                            worksheet.Cells[totalRow + 2, 11].Value = "Final Total";
                            worksheet.Cells[totalRow + 2, 11].Style.Font.Bold = true;
                            worksheet.Cells[totalRow + 2, 12].Value = finalTotal;
                            worksheet.Cells[totalRow + 2, 12].Style.Font.Bold = true;
                            worksheet.Cells[totalRow + 2, 12].Style.Numberformat.Format = "$#,##0.00";

                            // Save the file
                            package.Save();

                            if (MessageBox.Show("File created successfully! Do you want to open it?", "Success", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                            {
                                if (!string.IsNullOrEmpty(saveFilePath))
                                {
                                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                                    {
                                        FileName = saveFilePath,
                                        UseShellExecute = true
                                    });
                                }
                                else
                                {
                                    MessageBox.Show("The file path is not valid.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                            }
                        }
                        MessageBox.Show("File processed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<string[]> ParseCsvFile(string filePath)
        {
            var lines = new List<string[]>();
            using (var parser = new TextFieldParser(filePath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = true;

                while (!parser.EndOfData)
                {
                    var fields = parser.ReadFields();
                    if (fields != null)
                    {
                        lines.Add(fields);
                    }
                    else
                    {
                        MessageBox.Show("Encountered a null field while parsing the CSV file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            return lines;
        }

        private string GenerateVbaCode(int totalRow, int grandTotalColumn)
        {
            return $@"
Private Sub Workbook_Open()
    ' This event no longer adds buttons to avoid trust setting issues
    MsgBox ""Click the 'Create Gratuity Buttons' rectangle to add the gratuity buttons."", vbInformation, ""Instructions""
End Sub

Sub CreateButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(""Processed Data"")

    ' Add 20% Gratuity button
    Dim btn20 As Object
    Set btn20 = ws.Buttons.Add(400, 100, 100, 30) ' Left, Top, Width, Height
    btn20.OnAction = ""CalculateGratuity20""
    btn20.Caption = ""20% Gratuity""
    btn20.Font.Size = 10
    btn20.Font.Bold = True

    ' Add 30% Gratuity button
    Dim btn30 As Object
    Set btn30 = ws.Buttons.Add(400, 150, 100, 30) ' Left, Top, Width, Height
    btn30.OnAction = ""CalculateGratuity30""
    btn30.Caption = ""30% Gratuity""
    btn30.Font.Size = 10
    btn30.Font.Bold = True
End Sub

Sub CalculateGratuity20()
    Dim GrandTotal As Double
    Dim Gratuity As Double
    Dim FinalTotal As Double

    GrandTotal = Worksheets(""Processed Data"").Cells({totalRow}, {grandTotalColumn}).Value
    Gratuity = GrandTotal * 0.2
    FinalTotal = GrandTotal + Gratuity

    Worksheets(""Processed Data"").Cells({totalRow + 1}, {grandTotalColumn - 1}).Value = ""20% Gratuity""
    Worksheets(""Processed Data"").Cells({totalRow + 1}, {grandTotalColumn}).Value = Gratuity
    Worksheets(""Processed Data"").Cells({totalRow + 2}, {grandTotalColumn - 1}).Value = ""Grand Final Total""
    Worksheets(""Processed Data"").Cells({totalRow + 2}, {grandTotalColumn}).Value = FinalTotal
End Sub

Sub CalculateGratuity30()
    Dim GrandTotal As Double
    Dim Gratuity As Double
    Dim FinalTotal As Double

    GrandTotal = Worksheets(""Processed Data"").Cells({totalRow}, {grandTotalColumn}).Value
    Gratuity = GrandTotal * 0.3
    FinalTotal = GrandTotal + Gratuity

    Worksheets(""Processed Data"").Cells({totalRow + 1}, {grandTotalColumn - 1}).Value = ""30% Gratuity""
    Worksheets(""Processed Data"").Cells({totalRow + 1}, {grandTotalColumn}).Value = Gratuity
    Worksheets(""Processed Data"").Cells({totalRow + 2}, {grandTotalColumn - 1}).Value = ""Grand Final Total""
    Worksheets(""Processed Data"").Cells({totalRow + 2}, {grandTotalColumn}).Value = FinalTotal
End Sub";
        }
    }
}
