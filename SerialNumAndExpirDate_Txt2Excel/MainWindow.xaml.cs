﻿using SwiftExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace BarCodeDescrExpirDate_Txt2Excel
{
    public partial class MainWindow : Window
    {
        readonly Label StatusLabel;
        String FileName;
        public MainWindow()
        {
            InitializeComponent();
            StatusLabel = (Label)this.FindName("StatusLabel_");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Stream FileStream = SelectFile();

            if(FileStream != null)
            {
                StatusLabel.Content = "Status: lettura dati...";
                IList<String> contents = new List<String>();
                //Read the txt line by line
                using (var sr = new StreamReader(FileStream))
                {
                    while (!sr.EndOfStream)
                    {
                        try
                        {
                            String line = sr.ReadLine();
                            if(line != null)
                            {
                                contents.Add(line);
                            }
                        }
                        catch (Exception ex)
                        {
                            PrintError(ex);
                        }
                    }
                }
                //Convert the lines into a list of RowItem
                IList<RowItem> Rows = GetDatas(contents);
                StatusLabel.Content = "Status: inserimento dati (0%)...";
                //Create the excel file
                CreateExcelFile(Rows, FileName.Replace(".txt", ".xlsx"));
            }
        }

        private void CreateExcelFile(IList<RowItem> rows, String FileName)
        {
            try
            {
                //Create a table
                var sheet = new Sheet
                {
                    Name = "Prodotti",
                    ColumnsWidth = new List<double> { 10, 50, 70, 10 }
                };

                var ew = new ExcelWriter(FileName, sheet);

                //Header
                ew.Write("Codice a barre", 1, 1);
                ew.Write("Descrizione", 2, 1);
                ew.Write("Scadenza", 3, 1);

                int i = 2;
                //Populate cells
                foreach (RowItem row in rows)
                {
                    ew.Write(row.BarCode, 1, i);
                    ew.Write(row.Description, 2, i);
                    ew.Write(row.Expiration, 3, i);
                    int Percentage = ((i - 2) / rows.Count) * 100;
                    StatusLabel.Content = "Status: lettura dati (" + Percentage + "%)...";
                    i++;
                }

                ew.Save();

                StatusLabel.Content = "Status: file excel salvato nella stessa cartella del txt!";
            }
            catch (Exception ex)
            {
                PrintError(ex);
            }
        }

        private void PrintError(Exception ex)
        {
            String errorMessage = "Error: ";
            errorMessage = String.Concat(errorMessage, ex.Message);
            errorMessage = String.Concat(errorMessage, " Line: ");
            errorMessage = String.Concat(errorMessage, ex.Source);

            StatusLabel.Content = "Status: si è verificato un errore: " + errorMessage;
        }

        private static IList<RowItem> GetDatas(IList<string> contents)
        {
            IList<RowItem> Rows = new List<RowItem>(); 
            foreach(String Row in contents)
            {
                String[] Cols = Row.Split(";");
                //If the file is formatted correctly
                if(Cols.Length > 5)
                {
                    //If the expiration date is in the format DDMMYY, convert the month from int to string
                    if(Cols[4].Length == 6)
                    {
                        String Day = Cols[4].Substring(0, 2);
                        String Month = Cols[4].Substring(2, 2);
                        String Year = Cols[4].Substring(4, 2);

                        switch (Month)
                        {
                            case "01": Month = "Gennaio"; break;
                            case "02": Month = "Febbraio"; break;
                            case "03": Month = "Marzo"; break;
                            case "04": Month = "Aprile"; break;
                            case "05": Month = "Maggio"; break;
                            case "06": Month = "Giugno"; break;
                            case "07": Month = "Luglio"; break;
                            case "08": Month = "Agosto"; break;
                            case "09": Month = "Settembre"; break;
                            case "10": Month = "Ottobre"; break;
                            case "11": Month = "Novembre"; break;
                            case "12": Month = "Dicemebre"; break;
                        }

                        Rows.Add(new RowItem(Cols[0], Cols[2], Day + " " + Month + " " + Year));
                    }
                    else
                    {

                        Rows.Add(new RowItem(Cols[0], Cols[2], Cols[4]));
                    }
                }
            }

            return Rows;
        }

        private Stream SelectFile()
        {
            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new()
            {
                DefaultExt = ".txt", // Default file extension
                Filter = "Text documents (.txt)|*.txt" // Filter files by extension
            };

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                StatusLabel.Content = "Stato: elaborazione file " + dlg.FileName;
                this.FileName = dlg.FileName;
                return dlg.OpenFile();
            }
            else
            {
                StatusLabel.Content = "Operazione annullata/fallita";

                return null;
            }
        }
    }
}
