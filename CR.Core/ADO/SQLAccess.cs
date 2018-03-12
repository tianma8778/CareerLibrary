using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CR.Core.ADO
{

    public class SQLAccess : IDisposable
    {
        private SqlConnection conn = null;
        private SqlCommand cmd = null;

        private string connstr = ConfigurationManager.ConnectionStrings["default"].ConnectionString;

        private SQLAccess()
        { }

        public static SQLAccess GetInstance()
        {
            return new SQLAccess();
        }

        public static SQLAccess GetInstance(DBOption opt)
        {
            string constr = string.Empty;

            switch (opt)
            {
                case DBOption.Default:
                    constr = "default";
                    break;
                case DBOption.DiShang:
                    constr = "dishanscon";
                    break;
            }

            return GetInstance(constr);
        }

        public static SQLAccess GetInstance(string connstr)
        {
            return new SQLAccess() { connstr = ConfigurationManager.ConnectionStrings[connstr].ConnectionString };
        }

        #region 建立数据库连接对象
        /// <summary>
        /// 建立数据库连接
        /// </summary>
        /// <returns>返回一个数据库的连接SqlConnection对象</returns>
        public SqlConnection init()
        {
            try
            {
                conn = new SqlConnection(connstr);
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message.ToString());
            }
            return conn;
        }
        #endregion

        #region 设置SqlCommand对象
        /// <summary>
        /// 设置SqlCommand对象       
        /// </summary>
        /// <param name="cmd">SqlCommand对象 </param>
        /// <param name="cmdText">命令文本</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="cmdParms">参数集合</param>
        private void SetCommand(SqlCommand cmd, string cmdText, CommandType cmdType, SqlParameter[] cmdParms)
        {
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            cmd.CommandType = cmdType;
            if (cmdParms != null)
            {
                cmd.Parameters.AddRange(cmdParms);
            }
        }
        #endregion

        #region 执行相应的sql语句，返回相应的DataSet对象
        /// <summary>
        /// 执行相应的sql语句，返回相应的DataSet对象
        /// </summary>
        /// <param name="sqlstr">sql语句</param>
        /// <returns>返回相应的DataSet对象</returns>
        public DataSet GetDataSet(string sqlstr)
        {
            DataSet ds = new DataSet();
            try
            {
                init();
                SqlDataAdapter ada = new SqlDataAdapter(sqlstr, conn);
                ada.Fill(ds);
                conn.Close();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message.ToString());
            }
            return ds;
        }
        #endregion

        #region 执行相应的sql语句，返回相应的DataSet对象
        /// <summary>
        /// 执行相应的sql语句，返回相应的DataSet对象
        /// </summary>
        /// <param name="sqlstr">sql语句</param>
        /// <param name="tableName">表名</param>
        /// <returns>返回相应的DataSet对象</returns>
        public DataSet GetDataSet(string sqlstr, string tableName)
        {
            DataSet ds = new DataSet();
            try
            {
                init();
                SqlDataAdapter ada = new SqlDataAdapter(sqlstr, conn);
                ada.Fill(ds, tableName);
                conn.Close();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message.ToString());
            }
            return ds;
        }
        #endregion

        #region 执行不带参数sql语句，返回一个DataTable对象
        /// <summary>
        /// 执行不带参数sql语句，返回一个DataTable对象
        /// </summary>
        /// <param name="cmdText">相应的sql语句</param>
        /// <returns>返回一个DataTable对象</returns>
        public DataTable GetDataTable(string cmdText)
        {

            SqlDataReader reader;
            DataTable dt = new DataTable();
            try
            {
                init();
                cmd = new SqlCommand(cmdText, conn);
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                dt.Load(reader);
                reader.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
            return dt;
        }
        #endregion

        #region 执行带参数的sql语句或存储过程，返回一个DataTable对象
        /// <summary>
        /// 执行带参数的sql语句或存储过程，返回一个DataTable对象
        /// </summary>
        /// <param name="cmdText">sql语句或存储过程名</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="cmdParms">参数集合</param>
        /// <returns>返回一个DataTable对象</returns>
        public DataTable GetDataTable(string cmdText, CommandType cmdType, SqlParameter[] cmdParms)
        {
            SqlDataReader reader;
            DataTable dt = new DataTable();
            try
            {
                init();
                cmd = new SqlCommand();
                SetCommand(cmd, cmdText, cmdType, cmdParms);
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                dt.Load(reader);
                reader.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
            return dt;
        }

        public T GetEntity<T>(string cmdText, CommandType cmdType, params SqlParameter[] cmdParams) where T : new()
        {
            IList<T> ret = GetEntityList<T>(cmdText, cmdType, cmdParams);
            if (ret == null || ret.Count == 0)
                return default(T);
            return ret[0];
        }

        public IList<T> GetEntityList<T>(string cmdText, CommandType cmdType, params SqlParameter[] cmdParams) where T : new()
        {
            IList<T> lstObj = new List<T>();
            Type type = typeof(T);

            DataTable dt = GetDataTable(cmdText, cmdType, cmdParams);
            PropertyInfo[] propertyInfos = type.GetProperties();//获取指定类型里面的所有属性
            foreach (DataRow dr in dt.Rows)
            {
                T obj = new T();
                foreach (PropertyInfo propertyInfo in propertyInfos)
                {
                    object val = dr[propertyInfo.Name];
                    if (val != null && val != DBNull.Value)
                    {
                        if (val.GetType() == typeof(decimal) || val.GetType() == typeof(int))
                        {
                            propertyInfo.SetValue(obj, Convert.ToInt32(val), null);
                        }
                        else if (val.GetType() == typeof(DateTime))
                        {
                            propertyInfo.SetValue(obj, Convert.ToDateTime(val), null);
                        }
                        else if (val.GetType() == typeof(string))
                        {
                            propertyInfo.SetValue(obj, Convert.ToString(val), null);
                        }
                        else
                        {
                            propertyInfo.SetValue(obj, val, null);
                        }
                    }
                }
                lstObj.Add(obj);
            }

            return lstObj;
        }

        #endregion

        #region 执行不带参数sql语句，返回所影响的行数
        /// <summary>
        /// 执行不带参数sql语句，返回所影响的行数
        /// </summary>
        /// <param name="cmdText">增，删，改sql语句</param>
        /// <returns>返回所影响的行数</returns>
        public int ExecuteNonQuery(string cmdText)
        {
            int count;
            try
            {
                init();
                cmd = new SqlCommand(cmdText, conn);
                count = cmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
            return count;
        }
        #endregion

        #region 执行带参数sql语句或存储过程，返回所影响的行数
        /// <summary>
        /// 执行带参数sql语句或存储过程，返回所影响的行数
        /// </summary>
        /// <param name="cmdText">带参数的sql语句和存储过程名</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="cmdParms">参数集合</param>
        /// <returns>返回所影响的行数</returns>
        public int ExecuteNonQuery(string cmdText, CommandType cmdType, params SqlParameter[] cmdParms)
        {
            int count;
            try
            {
                init();
                cmd = new SqlCommand();
                SetCommand(cmd, cmdText, cmdType, cmdParms);
                count = cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                conn.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
            return count;
        }

        public string ExecuteNonQueryInBatch(ArrayList arr, Action<int, int> feedback)
        {
            StringBuilder errmsg = new StringBuilder();
            try
            {
                init();
                string cmdText;
                SqlParameter[] cmdParams;
                int i = 0;
                int arrcount = arr.Count;
                foreach (object[] item in arr)
                {
                    try
                    {
                        cmd = new SqlCommand();
                        cmdText = item[0].ToString();
                        cmdParams = item[1] as SqlParameter[];
                        SetCommand(cmd, cmdText, CommandType.Text, cmdParams);
                        cmd.ExecuteNonQuery();
                        if (feedback != null)
                        {
                            i++;
                            feedback(i, arrcount);
                        }
                    }
                    catch (Exception ex)
                    {
                        errmsg.Append(ex.Message + "/r/n");
                    }
                }
                cmd.Parameters.Clear();
                conn.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
            return errmsg.ToString();
        }
        #endregion

        #region 执行不带参数sql语句，返回一个从数据源读取数据的SqlDataReader对象
        /// <summary>
        /// 执行不带参数sql语句，返回一个从数据源读取数据的SqlDataReader对象
        /// </summary>
        /// <param name="cmdText">相应的sql语句</param>
        /// <returns>返回一个从数据源读取数据的SqlDataReader对象</returns>
        public SqlDataReader ExecuteReader(string cmdText)
        {
            SqlDataReader reader;
            try
            {
                init();
                cmd = new SqlCommand(cmdText, conn);
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
            return reader;
        }
        #endregion

        #region 执行带参数的sql语句或存储过程，返回一个从数据源读取数据的SqlDataReader对象
        /// <summary>
        /// 执行带参数的sql语句或存储过程，返回一个从数据源读取数据的SqlDataReader对象
        /// </summary>
        /// <param name="cmdText">sql语句或存储过程名</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="cmdParms">参数集合</param>
        /// <returns>返回一个从数据源读取数据的SqlDataReader对象</returns>
        public SqlDataReader ExecuteReader(string cmdText, CommandType cmdType, SqlParameter[] cmdParms)
        {
            SqlDataReader reader;
            try
            {
                init();
                cmd = new SqlCommand();
                SetCommand(cmd, cmdText, cmdType, cmdParms);
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
            return reader;
        }
        #endregion

        #region 执行不带参数sql语句,返回结果集首行首列的值object
        /// <summary>
        /// 执行不带参数sql语句,返回结果集首行首列的值object
        /// </summary>
        /// <param name="cmdText">相应的sql语句</param>
        /// <returns>返回结果集首行首列的值object</returns>
        public object ExecuteScalar(string cmdText)
        {
            object obj;
            try
            {
                init();
                cmd = new SqlCommand(cmdText, conn);
                obj = cmd.ExecuteScalar();
                conn.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
            return obj;
        }
        #endregion

        #region 执行带参数sql语句或存储过程,返回结果集首行首列的值object
        /// <summary>
        /// 执行带参数sql语句或存储过程,返回结果集首行首列的值object
        /// </summary>
        /// <param name="cmdText">sql语句或存储过程名</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="cmdParms">返回结果集首行首列的值object</param>
        /// <returns></returns>
        public object ExecuteScalar(string cmdText, CommandType cmdType, params SqlParameter[] cmdParms)
        {
            object obj;
            try
            {
                init();
                cmd = new SqlCommand();
                SetCommand(cmd, cmdText, cmdType, cmdParms);
                obj = cmd.ExecuteScalar();
                conn.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message.ToString());
            }
            return obj;
        }

        #endregion

        public void Dispose()
        {
            if (conn != null && conn.State != ConnectionState.Closed)
                conn.Close();
        }
    }

    public enum DBOption
    {
        Default,
        DiShang
    }
}
