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

        }

        private static Report Compare(ThreatTable threatTableOld, ThreatTable threatTableNew)
        {
            return null;
        }
    }
}
