using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VTMES3_RE.Common
{
    public class CamstarDatabase
    {
        private SqlConnection adoCon = new SqlConnection();
        private DataView returnView = null;
        public static String m_ConnectionString = string.Format("SERVER={0};Database={1};User ID={2};Password={3}"
                    , WrGlobal.Camstar_SQL_SERVER, WrGlobal.Camstar_SQL_Database, WrGlobal.Camstar_SQL_Id, WrGlobal.Camstar_SQL_Password);

        public CamstarDatabase()
        {
            adoCon.ConnectionString = m_ConnectionString;
        }

        ~CamstarDatabase()
        {
            this.Dispose(false);
        }

        #region IDisposable Support
        private bool disposedValue = false; // 중복 호출을 검색하려면

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: 관리되는 상태(관리되는 개체)를 삭제합니다.
                }

                // TODO: 관리되지 않는 리소스(관리되지 않는 개체)를 해제하고 아래의 종료자를 재정의합니다.
                // TODO: 큰 필드를 null로 설정합니다.
                DBClose();

                disposedValue = true;
            }
        }

        // TODO: 위의 Dispose(bool disposing)에 관리되지 않는 리소스를 해제하는 코드가 포함되어 있는 경우에만 종료자를 재정의합니다.
        // ~WrDatabase()
        // {
        //   // 이 코드를 변경하지 마세요. 위의 Dispose(bool disposing)에 정리 코드를 입력하세요.
        //   Dispose(false);
        // }

        // 삭제 가능한 패턴을 올바르게 구현하기 위해 추가된 코드입니다.
        public void Dispose()
        {
            // 이 코드를 변경하지 마세요. 위의 Dispose(bool disposing)에 정리 코드를 입력하세요.
            Dispose(true);
            // TODO: 위의 종료자가 재정의된 경우 다음 코드 줄의 주석 처리를 제거합니다.
            // GC.SuppressFinalize(this);
        }
        #endregion

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


        public SqlDataReader GetDataReader(string strQuery)
        {
            try
            {
                DBconnection();
                BindingSource bindsource = new BindingSource();
                SqlCommand cmd = new SqlCommand(strQuery, adoCon);
                return cmd.ExecuteReader();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            finally
            {
                DBClose();
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


        internal DataView GetTempProductPrice()
        {
            throw new NotImplementedException();
        }

        internal object GetDeptWarehouse(string p)
        {
            throw new NotImplementedException();
        }

        public bool ExecuteBulkCopy(DataTable dt)
        {
            bool result = false;

            DBconnection();
            SqlTransaction adoTran = adoCon.BeginTransaction();

            try
            {
                using (SqlBulkCopy bulk = new SqlBulkCopy(adoCon, SqlBulkCopyOptions.TableLock, adoTran))
                {
                    bulk.BulkCopyTimeout = 0;
                    bulk.BatchSize = dt.Rows.Count;
                    bulk.DestinationTableName = dt.TableName;

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        bulk.ColumnMappings.Add(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName);
                    }

                    bulk.WriteToServer(dt);
                    dt.Clear();

                    bulk.Close();
                }

                adoTran.Commit();

                result = true;
            }
            catch (Exception ex)
            {
                result = false;
                MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                adoTran.Dispose();
                DBClose();
            }

            return result;
        }

        public void WriteBulkInsert(DataTable dt)
        {
            SqlBulkCopy bulkCopy = new SqlBulkCopy(adoCon, SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.FireTriggers | SqlBulkCopyOptions.UseInternalTransaction, null);
            //  Insert 할 데이터베이스의 테이블 이름을 지정한다.
            bulkCopy.DestinationTableName = dt.TableName;
            adoCon.Open();
            bulkCopy.NotifyAfter = 1000;
            bulkCopy.BatchSize = 1000;

            foreach (DataColumn col in dt.Columns)
            {
                bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
            }

            bulkCopy.WriteToServer(dt);
            adoCon.Close();
        }

        public bool ExecuteQueryList(List<string> queryList)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = adoCon;
            bool blRtv = true;

            try
            {
                cmd.CommandTimeout = 180;
                cmd.Connection.Open();

                foreach (string query in queryList)
                {
                    cmd.CommandText = query;
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                blRtv = false;
                MessageBox.Show(ex.Message, "DB Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                cmd.Connection.Close();
            }

            return blRtv;
        }
    }
}
