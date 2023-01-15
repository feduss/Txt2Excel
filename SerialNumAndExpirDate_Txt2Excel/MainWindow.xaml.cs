using SwiftExcel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Linq;
using System.Windows.Media;

namespace BarCodeDescrExpirDate_Txt2Excel
{
    public partial class MainWindow : Window
    {
        private readonly TextBlock StatusLabel;
        private readonly TextBox ColumnsIndicesTB, ColumnsNamesTB, ValueSeparatorTB, DateColumnIndexTB;
        private List<int> columnsIndices = new List<int>();
        private List<String> columnsNames = new List<String>();
        private int dateColumnIndex;
        private String valueSeparator;
        private String InputPath;
        public MainWindow()
        {
            InitializeComponent();
            StatusLabel = (TextBlock)this.FindName("StatusLabel_");
            ColumnsIndicesTB = (TextBox)this.FindName("ColumnsIndicesTB_");
            ColumnsNamesTB = (TextBox)this.FindName("ColumnsNamesTB_");
            ValueSeparatorTB = (TextBox)this.FindName("ValueSeparatorTB_");
            DateColumnIndexTB = (TextBox)this.FindName("DateColumnIndexTB_");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Stream FileStream = SelectFile();

                if (FileStream != null)
                {
                    var Text = "Status: lettura dati...";
                    UpdateStatusLabel(Text, Brushes.Black);


                    List<String> contents = ReadLines(FileStream);

                    //Text = "Status: ordinamento dati...(" + contents.Count + ")";
                    //UpdateStatusLabel(Text);

                    if (dateColumnIndex != 0)
                    {
                        if (columnsIndices.Count() == 3 && columnsNames.Count() == 3)
                        {
                            List<RowWithDate> Rows, RowsNotParsable;
                            GetDatedDatas(contents, out Rows, out RowsNotParsable);
                            SortDatas(Rows);
                            Text = "Status: inserimento dati (0%)...";
                            UpdateStatusLabel(Text, Brushes.Black);
                            //Create the excel file
                            String DirName = "Scadenze prodotti";
                            String[] TempArray = InputPath.Split("\\");
                            String OutputPath = InputPath.Replace(TempArray[TempArray.Length - 1], "") + DirName;
                            if (!Directory.Exists(OutputPath))
                            {
                                Directory.CreateDirectory(OutputPath);
                            }
                            CreateDatedExcelFile(Rows, RowsNotParsable, OutputPath);
                        } else
                        {
                            Text = "Avendo inserito l'indice della colonna della data, puoi inserire al massimo 3 indici e titoli. Correggi e riprova.";
                            UpdateStatusLabel(Text, Brushes.Black);
                        }
                    } else
                    {

                        var datas = GetDatas(contents);

                        if (datas != null && datas.Count > 0)
                        {

                            String DirName = "Inventario";
                            String[] TempArray = InputPath.Split("\\");
                            String Filename = TempArray[TempArray.Length - 1];
                            String OutputPath = InputPath.Replace(Filename, "") + DirName + "\\" + Filename.Replace(".txt", "");
                            if (!Directory.Exists(OutputPath))
                            {
                                Directory.CreateDirectory(OutputPath);
                            }
                            CreateExcelFile(datas, OutputPath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PrintError(ex);
            }
        }

        private void UpdateStatusLabel(string Text, Brush TextColor)
        {
            StatusLabel.Text = Text;
            StatusLabel.Foreground = TextColor;
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
                var Text = "Stato: elaborazione file " + dlg.FileName;
                UpdateStatusLabel(Text, Brushes.Black);
                this.InputPath = dlg.FileName;
                return dlg.OpenFile();
            }
            else
            {
                var Text = "Operazione annullata/fallita";
                UpdateStatusLabel(Text, Brushes.Black);

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

        private List<Row> GetDatas(List<string> contents)
        {
            if (columnsIndices.Count <= 0)
            {
                var Text = "Devi inserire degli indici per continuare";
                UpdateStatusLabel(Text, Brushes.Red);
                return null;
            }
            else if (valueSeparator != null && !valueSeparator.Trim().Equals(""))
            {
                var Text = "Status: inserimento dati, attendere...";
                UpdateStatusLabel(Text, Brushes.Black);
                int i = 0;
                List<Row> rows = new List<Row> {};
                foreach (String Row in contents)
                {
                    String[] tempCols = Row.Split(valueSeparator);
                    String[] filteredCols = tempCols.Where((col, index) => columnsIndices.Contains(index)).ToArray<String>();

                    rows.Add(new Row(i, filteredCols));
                    i++;
                }
                return rows;
            } else
            {
                var Text = "Il separatore delle colonne non può essere vuoto";
                UpdateStatusLabel(Text, Brushes.Red);
                return null;
            }
        }

        private void CreateExcelFile(List<Row> Rows, String OutputPath)
        {
            try
            {
                ExcelWriter ew = null;
                int i = 1; //riga
                int j = 2; //colonna

                //Create a table
                Sheet newSheet = new Sheet
                {
                    Name = "Inventario",
                    ColumnsWidth = new List<double> { 10, 60, 70, 10 }
                };

                String OPath = OutputPath + ".xlsx";
                ew = new ExcelWriter(OPath, newSheet);

                //Header
                foreach (var name in columnsNames)
                {
                    //writw value colonna riga
                    ew.Write(name, j, i);
                    j++;
                }

                i = 2;
                j = 2;
                //Populate cells
                foreach (Row row in Rows)
                {
                    foreach(var value in row.Values)
                    {
                        //writw value colonna riga
                        ew.Write(value, j, i);
                        j++;
                    }

                    var Text = "Status: elaborazione riga " + (i - 1) + " su " + Rows.Count;
                    UpdateStatusLabel(Text, Brushes.Black);

                    j = 2;
                    i++;
                }

                if (ew != null)
                {
                    ew.Save();
                    var Text = "Status: file excel salvato/i nella cartella: " + OutputPath;
                    UpdateStatusLabel(Text, Brushes.Black);
                } else
                {
                    var Text = "Status: si è verificato un errore durante il salvaraggio del file.";
                    UpdateStatusLabel(Text, Brushes.Red);
                }
            }
            catch (Exception ex)
            {
                PrintError(ex);
            }
        }

        private void GetDatedDatas(List<string> contents, out List<RowWithDate> Rows, out List<RowWithDate> RowsNotParsable)
        {
            Rows = new List<RowWithDate>();
            RowsNotParsable = new List<RowWithDate>();
            CultureInfo cultureInfo = new CultureInfo("it-IT");
            var BarcodeIndex = columnsIndices.ToList<int>()[0];
            var DescriptionIndex = columnsIndices.ToList<int>()[1];
            int i = 0;
            foreach (String Row in contents)
            {
                String[] Cols = Row.Split(valueSeparator);
                //If the day has only month and year
                if (Cols[dateColumnIndex].Length == 4)
                {
                    int Month = Int32.Parse(Cols[dateColumnIndex].Substring(0, 2));
                    int Year = Int32.Parse("20" + Cols[dateColumnIndex].Substring(2, 2));
                    Cols[dateColumnIndex] = DateTime.DaysInMonth(Year, Month) + Cols[dateColumnIndex];
                }
                //If the day has only one digit, add a zero before
                //I can do this because i'm at 100% that the month has always 2 digits
                //because i know the structure of the file txt
                else if (Cols[dateColumnIndex].Length == 5)
                {
                    Cols[dateColumnIndex] = "0" + Cols[dateColumnIndex];
                }
                //If the expiration date is in the format DDMMYY, convert the month from int to string
                if (Cols[dateColumnIndex].Length == 6)
                {
                    String Day = Cols[dateColumnIndex].Substring(0, 2);
                    String Month = Cols[dateColumnIndex].Substring(2, 2);
                    String Year = Cols[dateColumnIndex].Substring(4, 2);
                    try
                    {
                        DateTime FormattedExpiration = DateTime.ParseExact(Cols[dateColumnIndex], "ddMMyy", cultureInfo);
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

                        Rows.Add(new RowWithDate(i, Cols[BarcodeIndex], Cols[DescriptionIndex], Day + " " + Month + " " + Year, FormattedExpiration, Month, Year));
                    }
                    //Errori di parsing
                    catch (FormatException ex)
                    {
                        RowsNotParsable.Add(new RowWithDate(i, Cols[BarcodeIndex], Cols[DescriptionIndex], Cols[dateColumnIndex], Month, Year));
                    }

                }
                //Errori di parsing
                else
                {
                    RowsNotParsable.Add(new RowWithDate(i, Cols[BarcodeIndex], Cols[DescriptionIndex], Cols[dateColumnIndex], "Error", "Error"));
                }
                i++;


            }
        }

        private List<RowWithDate> SortDatas(List<RowWithDate> RowItems)
        {
            RowItems.Sort();
            return RowItems;
        }

        //OLD
        private void CreateDatedExcelFile(List<RowWithDate> Rows, List<RowWithDate> RowsNotParsable, String OutputPath)
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
                foreach (RowWithDate row in Rows)
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
                        var Text1 = "Status: lettura dati con scadenza " + rowMonthYear + "...";
                        UpdateStatusLabel(Text1, Brushes.Black);
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
                foreach (RowWithDate row in RowsNotParsable)
                {
                    ew.Write(row.BarCode, 1, i);
                    ew.Write(row.Description, 2, i);
                    ew.Write(row.Expiration, 3, i);
                    //int Percentage = ((i - 2) / Rows.Count) * 100;
                    var Text1 = "Status: lettura dati con scadenze non valide...";
                    UpdateStatusLabel(Text1, Brushes.Black);

                    i++;
                }

                ew.Save();

                var Text = "Status: file excel salvato/i nella cartella: " + OutputPath;
                UpdateStatusLabel(Text, Brushes.Black);
            }
            catch (Exception ex)
            {
                PrintError(ex);
            }
        }

        private void WriteRow(List<RowWithDate> Rows, ExcelWriter ew, int i, int j, RowWithDate row)
        {
            ew.Write(row.BarCode, 1, i);
            ew.Write(row.Description, 2, i);
            ew.Write(row.Expiration, 3, i);
            int Percentage = ((j - 2) / Rows.Count) * 100;
            var Text = "Status: lettura dati (" + Percentage + "%)...";
            UpdateStatusLabel(Text, Brushes.Black);
        }

        private void PrintError(Exception ex)
        {
            String errorMessage = "Error: ";
            errorMessage = String.Concat(errorMessage, ex.Message);
            errorMessage = String.Concat(errorMessage, " Line: ");
            errorMessage = String.Concat(errorMessage, ex.Source);

            var Text = "Status: si è verificato un errore: \n\n" + errorMessage;
            UpdateStatusLabel(Text, Brushes.Red);
        }
        private void onColumnsIndicesTextChanged(object sender, TextChangedEventArgs args)
        {
            String[] indices = ColumnsIndicesTB.Text.Split(",");
            columnsIndices = new List<int>();
            if (indices.Count() > 1 && ColumnsIndicesTB.Text.Count() > 1) {
                var Text = "Stato: in attesa di un file.";
                UpdateStatusLabel(Text, Brushes.Black);
                foreach (var index in indices)
                {
                    try
                    {
                        var parsedIndex = Int32.Parse(index);
                        columnsIndices.Add(parsedIndex - 1);
                    }
                    catch (FormatException)
                    {
                        Text = "Controlla di aver inserito solo numeri tra ogni virgola.";
                        UpdateStatusLabel(Text, Brushes.Red);
                        break;
                    }

                }
            } else
            {
                var Text = "Controlla di aver separato gli indici con una virgola.";
                UpdateStatusLabel(Text, Brushes.Red);
            }
        }

        private void onColumnsNamesTextChanged(object sender, TextChangedEventArgs args)
        {
            String[] names = ColumnsNamesTB.Text.Split(",");
            columnsNames = new List<string>();
            if (names.Count() > 1 && ColumnsNamesTB.Text.Count() > 1)
            {
                var Text = "Stato: in attesa di un file.";
                UpdateStatusLabel(Text, Brushes.Black);
                foreach (var name in names)
                {
                    columnsNames.Add(name);

                }
            }
            else
            {
                var Text = "Controlla di aver separato i titoli con una virgola.";
                UpdateStatusLabel(Text, Brushes.Red);
            }
        }

        private void onValueSeparatorTextChanged(object sender, TextChangedEventArgs args)
        {
            valueSeparator = ValueSeparatorTB.Text;
        }

        private void onDateColumnTextChanged(object sender, TextChangedEventArgs args)
        {
            dateColumnIndex = Int32.Parse(DateColumnIndexTB.Text) - 1;
        }

    }

    
}
