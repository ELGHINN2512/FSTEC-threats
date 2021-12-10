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
                IExcelDataReader edr;
                FileStream stream = File.Open(this.fileName, FileMode.Open, FileAccess.Read);
                edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };
                dataSet = edr.AsDataSet(conf);
                edr.Close();
                stream.Close();
                this.dataTable = dataSet.Tables[0];
            }
            catch (FileNotFoundException ex)
            {
                MessageBoxResult result = MessageBox.Show(
                    "На данный момент у вас осутствует локальная база данных \n" +
                    "угроз безопасности информации с сайта ФСТЭК. Хотите загрузить?",
                    "Запрос на загрузку файла",
                    MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
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

        public void DownloadFile(string fileName)
        {
            WebClient webload = new WebClient();
            webload.DownloadFileCompleted += new AsyncCompletedEventHandler(Completed);
            webload.DownloadFileAsync(new Uri("https://bdu.fstec.ru/files/documents/thrlist.xlsx"), fileName);
        }

        private void Completed(object Sender, AsyncCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show("Не удалось скачать файл. Проверьте наличие интернет соединения");
            }
            else
            {
                MessageBox.Show("Cкачивание успешно завершено!\nВыберите режим отображения в меню сверху");
            }
        }

        public void Update()
        {
            if (this.dataTable == null)
            {
                MessageBox.Show("Для начала требуется скачать локальную базу данных.\n Выберите режим отображения в меню сверху.");
                if(File.Exists("thrlist.xlsx"))
                {
                    ReadFile();
                }
                else
                {
                    DownloadFile("thrlist.xlsx");
                    ReadFile();
                }
            }

            try
            {
                ThreatTable NewThreatTable = new ThreatTable("thrlist1.xlsx");
                if(!File.Exists("thrlist1.xlsx"))
                    NewThreatTable.DownloadFile(NewThreatTable.fileName);
                NewThreatTable.ReadFile();
                Report report = ThreatTable.Compare(this, NewThreatTable);
                MessageBox.Show(report.ToString());
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private static Report Compare(ThreatTable threatTableOld, ThreatTable threatTableNew)
        {
            Report report = new Report();
            DataTable dataTableNew = threatTableNew.dataTable;
            DataTable dataTableOld = threatTableOld.dataTable;
            if (threatTableOld == null | threatTableNew == null)
                return null;

            //Отслеживает новые уязвмости
            /*
            if(dataTableNew.Rows.Count != dataTableOld.Rows.Count)
            {
                for (int i = dataTableOld.Rows.Count; i < dataTableOld.Rows.Count; i++)
                {
                }
            }
            */
            
            //Отслеживаем изменения
            /*
            if(dataTableOld.Rows.Count <= dataTableNew.Rows.Count)
            {
                for (int i = 0; i < dataTableNew.Rows.Count; i++)
                {
                    if(dataTableNew.Rows[i][0] == dataTableOld.Rows[i][0])
                    {
                        for (int j = 1; i < 10; i++)
                        {
                            if (dataTableNew.Rows[i][j] != dataTableOld.Rows[i][j])
                            {
                                report.AddChange(dataTableNew.Rows[i][j].ToString(),
                                    dataTableNew.Columns[j].ColumnName,
                                    dataTableOld.Rows[i][j].ToString(),
                                    dataTableNew.Rows[i][j].ToString());
                            }
                        }
                    }
                }
            }
            else
            {
                for (int i = 0; i < dataTableOld.Rows.Count; i++)
                {
                    if (dataTableNew.Rows[i][0] == dataTableOld.Rows[i][0])
                    {
                        for (int j = 1; i < 10; i++)
                        {
                            if (dataTableNew.Rows[i][j] != dataTableOld.Rows[i][j])
                            {
                                report.AddChange(dataTableNew.Rows[i][j].ToString(),
                                    dataTableNew.Columns[j].ColumnName,
                                    dataTableOld.Rows[i][j].ToString(),
                                    dataTableNew.Rows[i][j].ToString());
                            }
                        }
                    }
                }
            }
            */
            return report;
        }
    }
}
