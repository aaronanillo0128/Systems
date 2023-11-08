using System;
using Microsoft.VisualBasic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

static class CRUD
{
    public static SqlConnection con = connection.mysqldb();
    public static SqlCommand cmd = new SqlCommand();
    public static SqlDataAdapter da = new SqlDataAdapter();
    public static DataTable dt = new DataTable();
    public static DataSet ds = new DataSet();
    public static int resultSQL;
    public static Image getIMG = null/* TODO Change to default(_) if this is not a reference type */;

    public static bool RETRIEVEDTG(DataGridView DTG, string SQL)
    {
        try
        {
            con.Open();
            cmd = new SqlCommand();
            {
                var withBlock = cmd;
                withBlock.Connection = con;
                withBlock.CommandText = SQL;
            }
            da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            dt = new DataTable();
            da.Fill(dt);

            {
                var withBlock = DTG;
                withBlock.DataSource = dt;
            }

            if (dt.Rows.Count > 0)
                return true;
            else
                return false;
        }


        catch (Exception ex)
        {

            return false;
        }
        finally
        {
            con.Close();
            da.Dispose();
        }
    }
    public static bool RETRIEVESINGLE(string SQL)
    {
        try
        {
            con.Open();
            cmd = new SqlCommand();
            {
                var withBlock = cmd;
                withBlock.Connection = con;
                withBlock.CommandText = SQL;
            }
            da = new SqlDataAdapter();
            da.SelectCommand = cmd;
            dt = new DataTable();
            da.Fill(dt);

            if (dt.Rows.Count > 0)
                return true;
            else
                return false;
        }
        catch (Exception ex)
        {

            return false;
        }
        finally
        {
            con.Close();
            da.Dispose();
        }
    }

    public static bool CUD(string sql)
    {
        try
        {
            con.Open();
            {
                var withBlock = cmd;
                withBlock.Connection = con;
                withBlock.CommandText = sql;
                resultSQL = cmd.ExecuteNonQuery();
            }
            if (resultSQL > 0)
                return true;
            else
                return false;
        }
        catch (Exception ex)
        {

            return false;
        }
        finally
        {
            con.Close();
            da.Dispose();
        }
    }

    public static void reloadtxt(string sql)
    {
        try
        {
            con.Open();
            {
                var withBlock = cmd;
                withBlock.Connection = con;
                withBlock.CommandText = sql;
            }
            dt = new DataTable();
            da = new SqlDataAdapter(sql, con);
            da.Fill(dt);
        }
        catch (Exception ex)
        {

        }
        con.Close();
        da.Dispose();
    }
}