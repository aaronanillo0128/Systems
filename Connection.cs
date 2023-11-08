
using System.Data.SqlClient;


class connection
{
    private static string conString = "Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd;"; //>>> Insert Constring here!

    public static SqlConnection mysqldb()
    {
        return new SqlConnection(conString);
    }
    public static SqlConnection con = mysqldb();
}



