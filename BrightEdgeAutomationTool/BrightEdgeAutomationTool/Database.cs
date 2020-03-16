using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Data.Sqlite;

namespace BrightEdgeAutomationTool
{
    #region MDF
    /*
    class Database
    {
        private SqlConnection Connection;
        public Database()
        {
            try
            {
                var myPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
                //var mdfPath = Path.Combine(myPath, "database.mdf");
                var mdfPath = "|DataDirectory|\\database.mdf";
                string sqlCon = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + mdfPath + ";Integrated Security=True";
                Connection = new SqlConnection(sqlCon);

                Connection.Open();
                //Console.WriteLine(mdfPath);
                //Console.WriteLine("Connection opened");
            }
            catch(Exception e)
            {
                MessageBox.Show($"Error while connecting to database: {e.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }


        }

        public User GetUser()
        {
            string query = "SELECT * FROM dbo.settings";

            SqlCommand cmd = new SqlCommand(query, Connection);

            var dataReader = cmd.ExecuteReader();

            while(dataReader.Read())
            {
                return new User { Email= dataReader.GetString(1), Password = dataReader.GetString(2) };
            }

            return null;
        }

        public User UpdateUser(string email, string password)
        {
            try
            {
                string query = $"UPDATE dbo.settings SET Email = '{email}', [Password]='{password}' WHERE Id IS NOT NULL";

                var command = new SqlCommand(query, Connection);
                SqlDataAdapter sqlDataAdap = new SqlDataAdapter(command);
                command.ExecuteNonQuery();
                command.Dispose();
                return new User { Email = email, Password = password};
            }
            catch (Exception e)
            {

            }
            

            return null;
        }

        public bool Close()
        {
            try
            {
                Connection.Close();
                return true;
            }
            catch (Exception e)
            {

            }

            return false;
        }
    }
    */
    #endregion

    /// <summary>
    /// User settings
    /// </summary>
    public class User
    {
        public string Email { get; set; }
        public string Password { get; set; }
        public bool RunBrightEdge { get; set; }
        public bool RunRankTracker { get; set; }
        public string RTExportPath { get; set; }
    }


    public class DbCreator
    {
        SQLiteConnection dbConnection;
        SQLiteCommand command;
        string sqlCommand;
        string dbPath = System.Environment.CurrentDirectory + "\\database";
        string dbFilePath;

        public DbCreator()
        {
            try
            {
                createDbFile();
                createDbConnection();
                createTables();
            }
            catch(Exception e)
            {
                MessageBox.Show($"Error while initializing to database: {e.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
        }

        public void createDbFile()
        {
            if (!string.IsNullOrEmpty(dbPath) && !Directory.Exists(dbPath))
                Directory.CreateDirectory(dbPath);
            dbFilePath = dbPath + "\\database.db";
            if (!System.IO.File.Exists(dbFilePath))
            {
                SQLiteConnection.CreateFile(dbFilePath);
            }
        }

        public string createDbConnection()
        {
            string strCon = string.Format("Data Source={0};", dbFilePath);
            dbConnection = new SQLiteConnection(strCon);
            dbConnection.Open();
            command = dbConnection.CreateCommand();
            return strCon;
        }

        public void createTables()
        {
            if (!checkIfExist("settings"))
            {
                sqlCommand = "CREATE TABLE settings ( id INTEGER PRIMARY KEY AUTOINCREMENT, email VARCHAR(62), password VARCHAR(32)," +
                    " run_bright_edge BOOL DEFAULT 1, run_rank_tracker BOOL DEFAULT 0, rt_export_path VARCHAR(260) DEFAULT '' )";
                executeQuery(sqlCommand);
                executeQuery("INSERT INTO settings (email, password) VALUES ('john.connolly@galileotechmedia.com', '')");
            }

        }

        public bool checkIfExist(string tableName)
        {
            command.CommandText = "SELECT name FROM sqlite_master WHERE name='" + tableName + "'";
            var result = command.ExecuteScalar();

            return result != null && result.ToString() == tableName ? true : false;
        }

        public int executeQuery(string sqlCommand)
        {
            SQLiteCommand triggerCommand = dbConnection.CreateCommand();
            triggerCommand.CommandText = sqlCommand;
            var result = triggerCommand.ExecuteNonQuery();
            return result;
        }

        public bool checkIfTableContainsData(string tableName)
        {
            command.CommandText = "SELECT count(*) FROM " + tableName;
            var result = command.ExecuteScalar();

            return Convert.ToInt32(result) > 0 ? true : false;
        }

        public User GetUser()
        {
            string statement = "SELECT * FROM settings LIMIT 1";

            var cmd = new SQLiteCommand(statement, dbConnection);
            SQLiteDataReader rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                return new User { Email = rdr.GetString(1), Password = rdr.GetString(2),
                    RunBrightEdge = rdr.GetBoolean(3), RunRankTracker = rdr.GetBoolean(4),
                    RTExportPath = rdr.GetString(5)
                };
            }

            return null;
        }

        public User UpdateUser(string email, string password, bool runBrightEdge, bool runRankTracker, string export_path)
        {
            try
            {
                var runBE = runBrightEdge ? 1 : 0;
                var runRT = runRankTracker ? 1 : 0;

                string query = $"UPDATE settings SET email = '{ email }', password='{ password }', " +
                    $"run_bright_edge = '{ runBE }', run_rank_tracker = '{ runRT }', rt_export_path = '{ export_path }' " +
                    $"WHERE Id IS NOT NULL";
                var result = executeQuery(query);
                if (result > 0)
                    return new User { Email = email, Password = password ,RTExportPath = export_path, RunBrightEdge = runBrightEdge, RunRankTracker = runRankTracker };
                else
                    return null;
            }
            catch (Exception e)
            {
            }

            return null;
        }

        public bool Close()
        {
            try
            {
                dbConnection.Close();
                return true;
            }
            catch (Exception e)
            {
            }
            return false;
        }


        public void fillTable()
        {
            if (!checkIfTableContainsData("MY_TABLE"))
            {
                sqlCommand = "insert into MY_TABLE (code_test_type) values (999)";
                executeQuery(sqlCommand);
            }
        }
    }

}
