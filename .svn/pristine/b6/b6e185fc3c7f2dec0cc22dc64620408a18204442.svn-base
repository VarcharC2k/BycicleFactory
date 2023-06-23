using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VTMES3_RE.Common
{
    public class Database2
    {
        private SqlConnection adoCon = new SqlConnection();
        private DataView returnView = null;
        public static String m_ConnectionString = "SERVER=RY-MESDB-SVR01,1435;Database=IFRY;User ID=sa;Password=Dentalimageno.1";   //데이터베이스 연결 스트링        

        public Database2()
        {
            adoCon.ConnectionString = m_ConnectionString;
        }
        public void DBconnection()
        {
            try
            {
                adoCon.Open();

                if (adoCon.State != ConnectionState.Open)
                {
                    MessageBox.Show("DB 연결 실패", "DB 연결", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            catch (Exception ex)
            {
                if (adoCon.State != ConnectionState.Open)
                {

                    MessageBox.Show(ex.Message, "DB 연결", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }//end try
        }//end function

        public void DBClose()
        {
            if (adoCon.State == ConnectionState.Open)
            {
                adoCon.Close();
            }

        }//end function

        public SqlCommand GetProcedure(string procedurename)
        {
            SqlCommand cmd = new SqlCommand(procedurename, adoCon);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = procedurename;
            return cmd;
        }
        public DataView GetDataView(string tableName, string strQuery)
        {
            DataSet ds = new DataSet();
            SqlDataAdapter SqlAdapter = new SqlDataAdapter(strQuery, adoCon);
            try
            {
                SqlAdapter.SelectCommand = new SqlCommand(strQuery, adoCon);
                SqlAdapter.SelectCommand.CommandTimeout = 180;
                SqlAdapter.Fill(ds, tableName);

                returnView = ds.Tables[tableName].DefaultView;
                return returnView;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
            }
        }//end function

        public DataView GetDataView2(string tableName, string strQuery)
        {
            DataSet ds = new DataSet();
            SqlDataAdapter SqlAdapter = new SqlDataAdapter(strQuery, adoCon);
            try
            {
                SqlAdapter.SelectCommand = new SqlCommand(strQuery, adoCon);
                SqlAdapter.SelectCommand.CommandTimeout = 180;
                SqlAdapter.Fill(ds, tableName);

                returnView = ds.Tables[tableName].DefaultView;
                return returnView;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
            }
        }//end function

        public DataSet GetDataSet(string tableName, string strQuery)
        {
            DataSet ds = new DataSet();
            SqlDataAdapter SqlAdapter = new SqlDataAdapter(strQuery, adoCon);
            try
            {

                SqlAdapter.SelectCommand = new SqlCommand(strQuery, adoCon);
                SqlAdapter.Fill(ds, tableName);
                return ds;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }
            finally
            {
            }

        }//end function

        public DataTable GetDataTable(string strQuery)
        {
            DataTable dt = new DataTable();
            SqlDataAdapter SqlAdapter = new SqlDataAdapter(strQuery, adoCon);
            try
            {
                SqlAdapter.SelectCommand = new SqlCommand(strQuery, adoCon);
                SqlAdapter.Fill(dt);
                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
            }

        }//end function

        public DataRowView GetDataRecord(string strQuery)
        {

            DataSet ds = new DataSet();
            SqlDataAdapter SqlAdapter = new SqlDataAdapter(strQuery, adoCon);

            try
            {
                SqlAdapter.SelectCommand = new SqlCommand(strQuery, adoCon);
                SqlAdapter.Fill(ds, "OneRow");

                returnView = ds.Tables["OneRow"].DefaultView;
                if (returnView != null && returnView.Count == 1)
                {
                    return returnView[0];
                }
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public bool ExecuteQuery(string strQuery)
        {
            SqlCommand cmd = new SqlCommand(strQuery, adoCon);
            bool blRtv = true;

            try
            {
                cmd.CommandTimeout = 180;
                cmd.Connection.Open();
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                blRtv = false;
                MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Connection.Close();
                //a.Dispose();
            }//end try

            return blRtv;

        }//end function

        public int ExecuteQueryAndReturnRows(string strQuery)
        {
            int rtv = 0;

            SqlCommand cmd = new SqlCommand(strQuery, adoCon);

            try
            {
                cmd.Connection.Open();
                rtv = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
            finally
            {
                cmd.Connection.Close();
                //a.Dispose();
            }//end try
            return rtv;
        }//end 

        public void WriteBulkInsert(DataTable dt)
        {
            ExecuteQuery("Truncate Table " + dt.TableName);

            SqlBulkCopy bulkCopy = new SqlBulkCopy(adoCon, SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.FireTriggers | SqlBulkCopyOptions.UseInternalTransaction, null);
            //  Insert 할 데이터베이스의 테이블 이름을 지정한다.
            bulkCopy.DestinationTableName = dt.TableName;
            adoCon.Open();
            bulkCopy.NotifyAfter = 1000;
            bulkCopy.BatchSize = 1000;
            bulkCopy.WriteToServer(dt);
            adoCon.Close();
        }

        internal DataView GetTempProductPrice()
        {
            throw new NotImplementedException();
        }

        internal object GetDeptWarehouse(string p)
        {
            throw new NotImplementedException();
        }

        public int ExecuteQueryList(List<string> queryList)
        {
            int count = 0;
            SqlCommand cmd = new SqlCommand();

            try
            {
                cmd.Connection = adoCon;
                cmd.CommandTimeout = 180;
                cmd.Connection.Open();

                foreach (string qry in queryList)
                {
                    cmd.CommandText = qry;
                    cmd.ExecuteNonQuery();

                    count++;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Connection.Close();
                //a.Dispose();
            }//end try

            return count;

        }//end function
    }
}
