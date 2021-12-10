using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Threading;


namespace FSTEC_threats
{
    class ThreatTable
    {
        private string fileName;

        private DataTable dataTable = null;
        public ThreatTable(string fileName)
        {
            this.fileName = fileName;
        }
        private void ReadFile()
        {
            DataSet dataSet = null;
            try
            {
                FileStream stream = File.Open(this.fileName, FileMode.Open, FileAccess.Read);
                IExcelDataReader edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };
                dataSet = edr.AsDataSet(conf);
                stream.Close();
                edr.Close();
                this.dataTable = dataSet.Tables[0];
            }
            catch (FileNotFoundException ex)
            {
                if (fileName == "thrlist.xlsx")
                {
                    MessageBoxResult result = MessageBox.Show(
                        "На данный момент у вас осутствует локальная база данных \n" +
                        "угроз безопасности информации с сайта ФСТЭК. Хотите загрузить?",
                        "Запрос на загрузку файла",
                        MessageBoxButton.YesNo);
                    if (result == MessageBoxResult.Yes)
                    {
                        DownloadFile(fileName);
                        ReadFile();
                    }
                }
                else
                {
                    DownloadFile(fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public DataView СreateDataViewWithFullInformation()
        {

            ReadFile();
            if (dataTable == null)
                return null;
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                if (dataTable.Rows[i][5].ToString() == "0")
                    dataTable.Rows[i][5] = "Нет";
                if (dataTable.Rows[i][5].ToString() == "1")
                    dataTable.Rows[i][5] = "Да";
                if (dataTable.Rows[i][6].ToString() == "0")
                    dataTable.Rows[i][6] = "Нет";
                if (dataTable.Rows[i][6].ToString() == "1")
                    dataTable.Rows[i][6] = "Да";
                if (dataTable.Rows[i][7].ToString() == "0")
                    dataTable.Rows[i][7] = "Нет";
                if (dataTable.Rows[i][7].ToString() == "1")
                    dataTable.Rows[i][7] = "Да";
            }
            return dataTable.AsDataView();
        }

        public DataView СreateTableWithAbbreviatedInformation()
        {
            ReadFile();
            if (dataTable == null)
                return null;
            DataTable newDataTable = new DataTable();
            DataColumn idColumn = new DataColumn("id", typeof(string));
            DataColumn nameColumn = new DataColumn("name", typeof(string));
            newDataTable.Columns.AddRange(new DataColumn[] { idColumn, nameColumn });
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                if (i != 0)
                {
                    DataRow row = newDataTable.NewRow();
                    row["id"] = "УБИ." + dataTable.Rows[i].ItemArray[0].ToString();
                    row["name"] = dataTable.Rows[i].ItemArray[1].ToString();
                    newDataTable.Rows.Add(row);
                }
                else
                {
                    DataRow row = newDataTable.NewRow();
                    row["id"] = dataTable.Rows[i].ItemArray[0].ToString();
                    row["name"] = dataTable.Rows[i].ItemArray[1].ToString();
                    newDataTable.Rows.Add(row);
                }
            }
            if (newDataTable != null)
                return newDataTable.AsDataView();
            else
                return null;
        }

        private void DownloadFile(string fileName)
        {
                WebClient webload = new WebClient();
                webload.DownloadFileCompleted += new AsyncCompletedEventHandler(Completed);
                webload.DownloadFile(new Uri("https://bdu.fstec.ru/files/documents/thrlist.xlsx"), fileName);
        }

        private void Completed(object Sender, AsyncCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show("Не удалось скачать файл. Проверьте наличие интернет соединения");
            }
            else
            {
                if(fileName == "thrlist.xlsx")
                    MessageBox.Show("Cкачивание успешно завершено!\nВыберите режим отображения в меню сверху");
            }
        }

        public void Update()
        {
            Report report = null;
            try
            {
                if (this.dataTable == null)
                {
                    MessageBox.Show("В начале требуется выбрать локальную базу угроз.\n Выберите режим отображения угроз в меню сверху.");
                    return;
                }
                ThreatTable threatTable = new ThreatTable("thrlist1.xlsx");
                if (!File.Exists("thrlist1.xlsx"))
                    threatTable.DownloadFile(threatTable.fileName);
                threatTable.ReadFile();
                threatTable.СreateDataViewWithFullInformation();
                report = ThreatTable.Compare(this, threatTable);
                File.Delete("thrlist.xlsx");
                File.Move("thrlist1.xlsx", "thrlist.xlsx");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            MessageBox.Show(report.ToString());

        }

        private static Report Compare(ThreatTable threatTableOld, ThreatTable threatTableNew)
        {
            Report report = new Report();
            DataTable dataTableNew = threatTableNew.dataTable;
            DataTable dataTableOld = threatTableOld.dataTable;
            if (threatTableOld == null | threatTableNew == null)
                return null;

            // Отслеживаем дополнения
            if(dataTableNew.Rows.Count > dataTableOld.Rows.Count)
            {
                for (int i = dataTableOld.Rows.Count; i < dataTableNew.Rows.Count; i++)
                {
                    report.AddNewThreat(dataTableNew.Rows[i][0].ToString(), dataTableNew.Rows[i][1].ToString());
                }
            }

            if (dataTableNew.Rows.Count < dataTableOld.Rows.Count)
            {
                for (int i = dataTableNew.Rows.Count; i < dataTableOld.Rows.Count; i++)
                {
                    report.AddNewThreat(dataTableOld.Rows[i][0].ToString(), dataTableOld.Rows[i][1].ToString());
                }
            }

            //Отслеживаем изменения
            if (dataTableOld.Rows.Count <= dataTableNew.Rows.Count)
            {
                for (int i = 0; i < dataTableOld.Rows.Count; i++)
                {
                    if (dataTableNew.Rows[i][0].ToString() == dataTableOld.Rows[i][0].ToString())
                    {
                        for (int j = 1; j < 10; j++)
                        {
                            if (dataTableNew.Rows[i][j].ToString() != dataTableOld.Rows[i][j].ToString())
                            {
                                report.AddChange(dataTableNew.Rows[i][0].ToString(),
                                    dataTableNew.Rows[0][j].ToString(),
                                    dataTableOld.Rows[i][j].ToString(),
                                    dataTableNew.Rows[i][j].ToString());
                            }
                        }
                    }
                }
            }
            else
            {
                for (int i = 0; i < dataTableNew.Rows.Count; i++)
                {
                    if (dataTableNew.Rows[i][0].ToString() == dataTableOld.Rows[i][0].ToString())
                    {
                        for (int j = 1; j < 10; j++)
                        {
                            if (dataTableNew.Rows[i][j].ToString() != dataTableOld.Rows[i][j].ToString())
                            {
                                report.AddChange(dataTableNew.Rows[1][0].ToString(),
                                    dataTableNew.Rows[0][j].ToString(),
                                    dataTableOld.Rows[i][j].ToString(),
                                    dataTableNew.Rows[i][j].ToString());
                            }
                        }
                    }
                }
            }
            return report;
        }
    }
}
