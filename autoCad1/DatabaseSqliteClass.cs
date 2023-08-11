using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SQLite;

namespace autoCad1
{
    public class DatabaseSqliteClass
    {

        string stringOfConnection = "Data Source=PartsDataBase.sqlite";

        //метод для получения информации из полей weight, FullNameTemplate, Diameter по конкретному ID
        public DataTable getInfoOfElememts(string guidOfElement)
        {
            string sqlSelect = "SELECT   replace(FullNameTemplate, '<$D$>', Diameter) as Name, Weight FROM 'Parts' where id='";
            var connection = new SQLiteConnection(stringOfConnection);
            connection.Open();
            SQLiteDataAdapter sQLiteDataAdapter = new SQLiteDataAdapter(sqlSelect + guidOfElement + "'", connection);
            DataTable dataTable = new DataTable();
            sQLiteDataAdapter.Fill(dataTable);
            return dataTable;

        }
    }
}
