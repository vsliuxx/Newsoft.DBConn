using System;
using System.Data;
using System.Data.SqlClient;

namespace Newsoft.DBConn
{
    /// <summary>
    /// 功能描述：通过SqlClient连接SqlServer数据库，实现查询与更新数据库操作
    /// 创建日期：2005年11月
    /// 创建作者：刘新学 lxx_@163.com
    /// 版本所有：东营市新视野软件有限公司
    /// </summary>
	public class ConnSqlServer
	{
        //定义变量	
        private string _connectString		= "" ;	//数据库连接字符串
        private string _errorMessage		= "" ;	//错误信息	

        public ConnSqlServer()
		{
			//
			// TODO: 在此处添加构造函数逻辑
			//
            //取得连接字符串OracleConnectionString
            _connectString = System.Configuration.ConfigurationManager.AppSettings["SqlServerConnectionString"];
		}

        /// <summary>
        /// 执行数据库查询操作
        /// </summary>
        /// <param name="_sql"></param>
        /// <returns></returns>
        public  DataSet executeQuery(string _sql)
        {
            //定义结果集
            DataSet ds = new DataSet();	

            //SqlClient连接数据
            SqlConnection Conn= new SqlConnection(_connectString);

            //打开数据库连接
            Conn.Open();
            
            //建立数据适配器
            SqlDataAdapter objAdp= new SqlDataAdapter(_sql,Conn);
            
	
            try
            {
                //建立连接		
                objAdp.Fill(ds); 								
            } 
            catch (Exception ex) 
            {	
                _errorMessage = ex.Message;
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
        public void executeUpdate(string _sql)
        {			
            //建立数据库连接
            SqlConnection Conn = new SqlConnection(_connectString);		
            
            //打开数据库连接
            Conn.Open();

            //建立数据库命令对象
            SqlCommand objCmd = new SqlCommand(_sql,Conn);

            try
            {
                //执行数据库更新操作
                objCmd.ExecuteNonQuery();
            }
            catch(Exception ex)
            {
                _errorMessage = ex.Message;
            }
            finally
            {
                objCmd.Dispose();
                Conn.Close();
                Conn.Dispose();
            }
        }

		
        //设置连接字符串中的信息
        public string ConnStr
        {
            set{_connectString = value;}
            get{return _connectString;}
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
