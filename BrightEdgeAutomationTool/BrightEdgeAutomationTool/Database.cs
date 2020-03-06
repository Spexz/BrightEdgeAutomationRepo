using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace BrightEdgeAutomationTool
{
    class Database
    {
        private SqlConnection Connection;
        public Database()
        {
            var myPath = System.IO.Path.GetDirectoryName( System.Reflection.Assembly.GetEntryAssembly().Location);
            var mdfPath = Path.Combine(myPath, "database.mdf");
            string sqlCon = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename="+ mdfPath +";Integrated Security=True";
            Connection = new SqlConnection(sqlCon);

            Connection.Open();
            //Console.WriteLine(mdfPath);
            //Console.WriteLine("Connection opened");
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

    public class User
    {
        public string Email { get; set; }
        public string Password { get; set; }
    }
}
