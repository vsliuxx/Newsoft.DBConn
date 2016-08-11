using System;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using Microsoft.Win32;

namespace Newsoft.DBConn
{
	/// <summary>
	/// 功能描述：通过OleDb连接Oracle数据库，实现查询与更新数据库操作
	/// 创建日期：2005年11月
	/// 版本所有：东营市新视野软件有限公司
	/// </summary>
	public class ConnOracle //: IDisposable

	{
		// 数据库连接字符串
		private string _connectString		= "" ;	

        // 错误信息	
		private string _errorMessage		= "" ;		
	
		/// <summary>
		/// 连接数据库
		/// </summary>
		public ConnOracle()
		{
			//
			// TODO: 在此处添加构造函数逻辑
			//	
	
			//取得连接字符串OracleConnectionString
            _connectString = System.Configuration.ConfigurationManager.AppSettings["OracleConnectionString"];

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

        public void Dispose()
        {
            GC.SuppressFinalize(true);
        }

		
		/// <summary>
        /// 设置连接字符串中的信息
		/// </summary>
		public string ConnectString
		{
            set
            {
                _connectString = value;
            }

            get
            {
                return _connectString;
            }
		}

        /// <summary>
        /// 获取错误信息 
        /// </summary>
        public string ErrorMessage
        {
            get
            {
                return this._errorMessage;
            }
        }

        /// <summary>
		/// 执行数据库查询操作
		/// </summary>
		/// <param name="_strSql"></param>
		/// <returns></returns>
		public  DataSet executeQuery(string _strSql)
		{
			//定义结果集
			DataSet ds = new DataSet();	

            //OLEDB连接数据
            OleDbConnection Conn= new OleDbConnection(_connectString);
            OleDbDataAdapter objAdp= new OleDbDataAdapter(_strSql,Conn);
	
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

                if (Conn.State == ConnectionState.Open)
                {
                    Conn.Close();
                    Conn.Dispose();
                }
			}
			return ds;		
		}
        
		/// <summary>
		/// 执行数据库更新操作
		/// </summary>
		/// <param name="_strSql">SQL语句</param>
		public void executeUpdate(string _strSql)
		{			
           //建立数据库连接
            OleDbConnection Conn = new OleDbConnection(_connectString);		

            // 打开数据库连接
            Conn.Open();

            //建立数据库命令对象
            OleDbCommand objCmd = new OleDbCommand(_strSql,Conn);

			try
			{
				//执行数据库更新操作
				objCmd.ExecuteNonQuery();
			}
			catch(Exception ex)
			{
                _errorMessage = ex.Message;
                throw new Exception("数据库更新操作异常：" + ex.Message);
			}
			finally
			{
				objCmd.Dispose();

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
        /// <param name="_strSql">查询语句</param>
        /// <returns></returns>
        public byte[] GetBlob(string _strSql)
        {
            // 存放Blob内容
            byte[] blob = new byte[0];

            //OLEDB连接数据
            OleDbConnection Conn = new OleDbConnection(_connectString);
            OleDbDataAdapter objAdp = new OleDbDataAdapter(_strSql, Conn);

            try
            {
                //定义结果集
                DataSet ds = new DataSet();

                //建立连接		
                objAdp.Fill(ds);

                DataRow dr = ds.Tables[0].Rows[0];

                if (!dr[0].ToString().Equals(""))
                {
                    blob = (byte[])dr[0];
                }
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                throw new Exception("数据库读取BLOB字段操作异常：" + ex.Message);
            }
            finally
            {
                objAdp.Dispose();
                Conn.Close();
                Conn.Dispose();
            }

            return blob;
        }
	}

}
