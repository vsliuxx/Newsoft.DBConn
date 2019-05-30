using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.Odbc;

/// <summary>
/// 功能描述：通过ConnOdbc连接数据库，实现查询与更新数据库操作
/// 创建日期：2005年11月
/// 创建作者：刘新学 lxx_@163.com
/// 版本所有：东营市新视野软件有限公司
/// </summary>
namespace Newsoft.DBConn
{
    public class ConnOdbc
    {
        //定义变量	
        private string _connectString = "";	//数据库连接字符串
        private string _errorMessage = "";	//错误信息	

        public ConnOdbc()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
            //取得连接字符串OdbcConnectionString
            _connectString = System.Configuration.ConfigurationManager.AppSettings["OdbcConnectionString"];

            //Odbc连接数据
            OdbcConnection Conn = new OdbcConnection(_connectString);

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
        public ConnOdbc(string keyName)
        {
            // 取得数据库连接
            _connectString = System.Configuration.ConfigurationManager.AppSettings[keyName];

            //Odbc连接数据
            OdbcConnection Conn = new OdbcConnection(_connectString);

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
            OdbcConnection Conn = new OdbcConnection(_connectString);

            //打开数据库连接
            Conn.Open();

            //建立数据适配器
            OdbcDataAdapter objAdp = new OdbcDataAdapter(_strSql, Conn);

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
            OdbcConnection Conn = new OdbcConnection(_connectString);

            //打开数据库连接
            Conn.Open();

            //建立数据库命令对象
            OdbcCommand objCmd = new OdbcCommand(_strSql, Conn);

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
