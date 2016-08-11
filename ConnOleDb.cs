using System;
using System.Data;
using System.Data.OleDb;

namespace Newsoft.DBConn
{
    /// <summary>
    /// 功能描述：通过OleDb连接SqlServer数据库，实现查询与更新数据库操作
    /// 创建日期：2005年11月
    /// 创建作者：刘新学 lxx_@163.com
    /// 版本所有：东营市新视野软件有限公司
    /// </summary>
    public class ConnOleDb
    {
        //定义变量	
        private string _connectString = "";	//数据库连接字符串
        private string _errorMessage = "";	//错误信息	

        public ConnOleDb()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
            //取得连接字符串OracleConnectionString
            _connectString = System.Configuration.ConfigurationManager.AppSettings["OleDbConnectionString"];

            //OLEDB连接数据
            OleDbConnection Conn = new OleDbConnection(_connectString);

            try
            {
                Conn.Open();
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                throw new Exception("数据库连接异常：" + ex.Message);
            }
            finally
            {
                if (Conn.State == ConnectionState.Open)
                {
                    Conn.Close();
                    Conn.Dispose();
                }
            }
        }

        /// <summary>
        /// 根据不同数据连接配置进行初始化
        /// </summary>
        /// <param name="keyName"></param>
        public ConnOleDb(string keyName)
        {
            // 取得数据库连接
            _connectString = System.Configuration.ConfigurationManager.AppSettings[keyName];

            //OLEDB连接数据
            OleDbConnection Conn = new OleDbConnection(_connectString);

            try
            {
                Conn.Open();
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                throw new Exception("数据库连接异常：" + ex.Message);
            }
            finally
            {
                if (Conn.State == ConnectionState.Open)
                {
                    Conn.Close();
                    Conn.Dispose();
                }
            }
        }
      

        /// <summary>
        /// 执行数据库查询操作
        /// </summary>
        /// <param name="_sql"></param>
        /// <returns></returns>
        public DataSet executeQuery(string _strSql)
        {
            //定义结果集
            DataSet ds = new DataSet();

            //SqlClient连接数据
            OleDbConnection Conn = new OleDbConnection(_connectString);

            //打开数据库连接
            Conn.Open();

            //建立数据适配器
            OleDbDataAdapter objAdp = new OleDbDataAdapter(_strSql, Conn);

            try
            {
                //建立连接		
                objAdp.Fill(ds);
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                throw new Exception("数据库查询操作异常：" + ex.Message);
            }
            finally
            {
                objAdp.Dispose();
                Conn.Close();
                Conn.Dispose();
            }
            return ds;
        }


        /// <summary>
        /// 执行数据库更新操作
        /// </summary>
        /// <param name="_sql"></param>
        public void executeUpdate(string _strSql)
        {
            //建立数据库连接
            OleDbConnection Conn = new OleDbConnection(_connectString);

            //打开数据库连接
            Conn.Open();

            //建立数据库命令对象
            OleDbCommand objCmd = new OleDbCommand(_strSql, Conn);

            try
            {
                //执行数据库更新操作
                objCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                throw new Exception("数据库更新操作异常：" + ex.Message);
            }
            finally
            {
                objCmd.Dispose();
                Conn.Close();
                Conn.Dispose();
            }
        }


        //设置连接字符串中的信息
        public string ConnectionString
        {
            set { _connectString = value; }
            get { return _connectString; }
        }

        /// <summary>
        /// 获取错误信息
        /// </summary>
        public string ErrorMessage
        {
            set { this._errorMessage = value; }
            get { return this._errorMessage; }
        }
    }
}
