using System;
using System.Data;
using System.Data.SqlClient;

namespace Newsoft.DBConn
{
    /// <summary>
    /// ����������ͨ��SqlClient����SqlServer���ݿ⣬ʵ�ֲ�ѯ��������ݿ����
    /// �������ڣ�2005��11��
    /// �������ߣ�����ѧ lxx_@163.com
    /// �汾���У���Ӫ������Ұ������޹�˾
    /// </summary>
	public class ConnSqlServer
	{
        //�������	
        private string _connectString		= "" ;	//���ݿ������ַ���
        private string _errorMessage		= "" ;	//������Ϣ	

        public ConnSqlServer()
		{
			//
			// TODO: �ڴ˴���ӹ��캯���߼�
			//
            //ȡ�������ַ���OracleConnectionString
            _connectString = System.Configuration.ConfigurationManager.AppSettings["SqlServerConnectionString"];
		}

        /// <summary>
        /// ִ�����ݿ��ѯ����
        /// </summary>
        /// <param name="_sql"></param>
        /// <returns></returns>
        public  DataSet executeQuery(string _sql)
        {
            //��������
            DataSet ds = new DataSet();	

            //SqlClient��������
            SqlConnection Conn= new SqlConnection(_connectString);

            //�����ݿ�����
            Conn.Open();
            
            //��������������
            SqlDataAdapter objAdp= new SqlDataAdapter(_sql,Conn);
            
	
            try
            {
                //��������		
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
        /// ִ�����ݿ���²���
        /// </summary>
        /// <param name="_sql"></param>
        public void executeUpdate(string _sql)
        {			
            //�������ݿ�����
            SqlConnection Conn = new SqlConnection(_connectString);		
            
            //�����ݿ�����
            Conn.Open();

            //�������ݿ��������
            SqlCommand objCmd = new SqlCommand(_sql,Conn);

            try
            {
                //ִ�����ݿ���²���
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

		
        //���������ַ����е���Ϣ
        public string ConnStr
        {
            set{_connectString = value;}
            get{return _connectString;}
        }

        /// <summary>
        /// ��ȡ������Ϣ
        /// </summary>
        public string ErrorMessage
        {
            set { this._errorMessage = value; }
            get { return this._errorMessage; }
        }
	}
}
