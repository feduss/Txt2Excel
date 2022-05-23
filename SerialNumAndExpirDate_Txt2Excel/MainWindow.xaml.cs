using SwiftExcel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace BarCodeDescrExpirDate_Txt2Excel
{
    public partial class MainWindow : Window
    {
        private readonly TextBlock StatusLabel;
        private String InputPath;
        public MainWindow()
        {
            InitializeComponent();
            StatusLabel = (TextBlock)this.FindName("StatusLabel_");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Stream FileStream = SelectFile();

                if (FileStream != null)
                {
                    StatusLabel.Text = "Status: lettura dati...";
                    new Thread(() =>
                    {
                        Thread.CurrentThread.IsBackground = true;

                        List<String> contents = ReadLines(FileStream);
                        //Convert the lines into a list of RowItem and sort them
                        StatusLabel.Dispatcher.Invoke(() => {
                            StatusLabel.Text = "Status: ordinamento dati...";
                        });
                        List<RowItem> Rows, RowsNotParsable;
                        GetDatas(contents, out Rows, out RowsNotParsable);
                        SortDatas(Rows);
                        StatusLabel.Dispatcher.Invoke(() => {
                            StatusLabel.Text = "Status: inserimento dati (0%)...";
                        });
                        //Create the excel file
                        String BaseFileName = "Scadenze prodotti";
                        String[] TempArray = InputPath.Split("\\");
                        String OutputPath = InputPath.Replace(TempArray[TempArray.Length - 1], "") + BaseFileName;
                        if (!Directory.Exists(OutputPath))
                        {
                            Directory.CreateDirectory(OutputPath);
                        }
                        CreateExcelFile(Rows, RowsNotParsable, OutputPath);
                    }).Start();
                }
            }
            catch (Exception ex)
            {
                PrintError(ex);
            }
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
                StatusLabel.Text = "Stato: elaborazione file " + dlg.FileName;
                this.InputPath = dlg.FileName;
                return dlg.OpenFile();
            }
            else
            {
                StatusLabel.Text = "Operazione annullata/fallita";

                return null;
            }
        }

        private List<String> ReadLines(Stream FileStream)
        {
            List<String> contents = new List<string>();
            var sr = new StreamReader(FileStream);
            while (!sr.EndOfStream)
            {
                try
                {
                    String line = sr.ReadLine();
                    if (line != null)
                    {
                        contents.Add(line);
                    }
                }
                catch (Exception ex)
                {
                    PrintError(ex);
                }
            }
            return contents;
        }

        private static void GetDatas(List<string> contents, out List<RowItem> Rows, out List<RowItem> RowsNotParsable)
        {
            Rows = new List<RowItem>();
            RowsNotParsable = new List<RowItem>();
            CultureInfo cultureInfo = new CultureInfo("it-IT");
            int i = 0;
            foreach (String Row in contents)
            {
                String[] Cols = Row.Split(";");
                //If the file is formatted correctly
                if (Cols.Length > 5)
                {
                    //If the day has only month and year
                    if (Cols[4].Length == 4)
                    {
                        int Month = Int32.Parse(Cols[4].Substring(0, 2));
                        int Year = Int32.Parse("20" + Cols[4].Substring(2, 2));
                        Cols[4] = DateTime.DaysInMonth(Year, Month) + Cols[4];
                    }
                    //If the day has only one digit, add a zero before
                    //I can do this because i'm at 100% that the month has always 2 digits
                    //because i know the structure of the file txt
                    else if (Cols[4].Length == 5)
                    {
                        Cols[4] = "0" + Cols[4];
                    }
                    //If the expiration date is in the format DDMMYY, convert the month from int to string
                    if (Cols[4].Length == 6)
                    {
                        String Day = Cols[4].Substring(0, 2);
                        String Month = Cols[4].Substring(2, 2);
                        String Year = Cols[4].Substring(4, 2);
                        try
                        {
                            DateTime FormattedExpiration = DateTime.ParseExact(Cols[4], "ddMMyy", cultureInfo);
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
                                case "12": Month = "Dicembre"; break;
                            }

                            Rows.Add(new RowItem(i, Cols[0], Cols[2], Day + " " + Month + " " + Year, FormattedExpiration, Month, Year));
                        }
                        //Errori di parsing
                        catch (FormatException ex)
                        {
                            RowsNotParsable.Add(new RowItem(i, Cols[0], Cols[2], Cols[4], Month, Year));
                        }

                    }
                    //Errori di parsing
                    else
                    {
                        RowsNotParsable.Add(new RowItem(i, Cols[0], Cols[2], Cols[4], "Error", "Error"));
                    }
                }
                i++;


            }
        }

        private List<RowItem> SortDatas(List<RowItem> RowItems)
        {
            RowItems.Sort();
            return RowItems;
        }

        private void CreateExcelFile(List<RowItem> Rows, List<RowItem> RowsNotParsable, String OutputPath)
        {
            //I can't create multiple sheet in one excel file because it requires SwiftExcelPro
            //So, i create multiple excel file
            try
            {
                ExcelWriter ew = null;
                int i = 2;
                int j = 2;
                String prevRowMonthYear = null;
                //Populate cells
                foreach (RowItem row in Rows)
                {
                    var rowMonthYear = row.Month + "" + row.Year;
                    //Create a new sheet if month and year change
                    if (prevRowMonthYear == null || !rowMonthYear.Equals(prevRowMonthYear))
                    {
                        if(ew != null)
                        {
                            ew.Save();
                        }
                        //Create a table
                        Sheet newSheet = new Sheet
                        {
                            Name = rowMonthYear,
                            ColumnsWidth = new List<double> { 10, 60, 70, 10 }
                        };

                        String OPath = OutputPath + "\\" + rowMonthYear + ".xlsx";
                        ew = new ExcelWriter(OPath, newSheet);

                        //Header
                        ew.Write("Codice a barre", 1, 1);
                        ew.Write("Descrizione", 2, 1);
                        ew.Write("Scadenza", 3, 1);

                        i = 2;
                    }
                    
                    if (ew != null)
                    {
                        prevRowMonthYear = rowMonthYear;
                        StatusLabel.Dispatcher.Invoke(() => {
                            StatusLabel.Text = "Status: lettura dati con scadenza " + rowMonthYear + "...";
                        });
                        WriteRow(Rows, ew, i, j, row);
                    }

                    j++;

                    i++;
                }

                if (ew != null)
                {
                    ew.Save();
                }

                //Create a table
                Sheet errorSheet = new Sheet
                {
                    Name = "Errori",
                    ColumnsWidth = new List<double> { 10, 60, 70, 10 }
                };

                String Path = OutputPath + "\\" + "DatiNonValidi.xlsx";
                ew = new ExcelWriter(Path, errorSheet);

                //Header
                ew.Write("Codice a barre", 1, 1);
                ew.Write("Descrizione", 2, 1);
                ew.Write("Scadenza", 3, 1);

                i = 2;

                //Populate cells with not parsable datas
                foreach (RowItem row in RowsNotParsable)
                {
                    ew.Write(row.BarCode, 1, i);
                    ew.Write(row.Description, 2, i);
                    ew.Write(row.Expiration, 3, i);
                    //int Percentage = ((i - 2) / Rows.Count) * 100;
                    StatusLabel.Dispatcher.Invoke(() => {
                        StatusLabel.Text = "Status: lettura dati con scadenze non valide...";
                    });

                    i++;
                }

                ew.Save();

                StatusLabel.Dispatcher.Invoke(() =>
                {
                    StatusLabel.Text = "Status: file excel salvato/i nella cartella: " + OutputPath;
                });
            }
            catch (Exception ex)
            {
                PrintError(ex);
            }
        }

        private void WriteRow(List<RowItem> Rows, ExcelWriter ew, int i, int j, RowItem row)
        {
            ew.Write(row.BarCode, 1, i);
            ew.Write(row.Description, 2, i);
            ew.Write(row.Expiration, 3, i);
            int Percentage = ((j - 2) / Rows.Count) * 100;
            StatusLabel.Dispatcher.Invoke(() =>
            {
                StatusLabel.Text = "Status: lettura dati (" + Percentage + "%)...";
            });
        }

        private void PrintError(Exception ex)
        {
            String errorMessage = "Error: ";
            errorMessage = String.Concat(errorMessage, ex.Message);
            errorMessage = String.Concat(errorMessage, " Line: ");
            errorMessage = String.Concat(errorMessage, ex.Source);

            StatusLabel.Dispatcher.Invoke(() =>
            {
                StatusLabel.Text = "Status: si è verificato un errore: \n\n" + errorMessage;
                //MessageBox.Show(errorMessage);
            });
        }

    }
}
