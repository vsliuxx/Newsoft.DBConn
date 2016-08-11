using System;
using System.Data;
using System.Data.OleDb;

namespace Newsoft.DBConn
{
    /// <summary>
    /// ����������ͨ��OleDb����SqlServer���ݿ⣬ʵ�ֲ�ѯ��������ݿ����
    /// �������ڣ�2005��11��
    /// �������ߣ�����ѧ lxx_@163.com
    /// �汾���У���Ӫ������Ұ������޹�˾
    /// </summary>
    public class ConnOleDb
    {
        //�������	
        private string _connectString = "";	//���ݿ������ַ���
        private string _errorMessage = "";	//������Ϣ	

        public ConnOleDb()
        {
            //
            // TODO: �ڴ˴���ӹ��캯���߼�
            //
            //ȡ�������ַ���OracleConnectionString
            _connectString = System.Configuration.ConfigurationManager.AppSettings["OleDbConnectionString"];

            //OLEDB��������
            OleDbConnection Conn = new OleDbConnection(_connectString);

            try
            {
                Conn.Open();
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                throw new Exception("���ݿ������쳣��" + ex.Message);
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
        /// ���ݲ�ͬ�����������ý��г�ʼ��
        /// </summary>
        /// <param name="keyName"></param>
        public ConnOleDb(string keyName)
        {
            // ȡ�����ݿ�����
            _connectString = System.Configuration.ConfigurationManager.AppSettings[keyName];

            //OLEDB��������
            OleDbConnection Conn = new OleDbConnection(_connectString);

            try
            {
                Conn.Open();
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                throw new Exception("���ݿ������쳣��" + ex.Message);
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
        /// ִ�����ݿ��ѯ����
        /// </summary>
        /// <param name="_sql"></param>
        /// <returns></returns>
        public DataSet executeQuery(string _strSql)
        {
            //��������
            DataSet ds = new DataSet();

            //SqlClient��������
            OleDbConnection Conn = new OleDbConnection(_connectString);

            //�����ݿ�����
            Conn.Open();

            //��������������
            OleDbDataAdapter objAdp = new OleDbDataAdapter(_strSql, Conn);

            try
            {
                //��������		
                objAdp.Fill(ds);
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                throw new Exception("���ݿ��ѯ�����쳣��" + ex.Message);
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
        public void executeUpdate(string _strSql)
        {
            //�������ݿ�����
            OleDbConnection Conn = new OleDbConnection(_connectString);

            //�����ݿ�����
            Conn.Open();

            //�������ݿ��������
            OleDbCommand objCmd = new OleDbCommand(_strSql, Conn);

            try
            {
                //ִ�����ݿ���²���
                objCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                _errorMessage = ex.Message;
                throw new Exception("���ݿ���²����쳣��" + ex.Message);
            }
            finally
            {
                objCmd.Dispose();
                Conn.Close();
                Conn.Dispose();
            }
        }


        //���������ַ����е���Ϣ
        public string ConnectionString
        {
            set { _connectString = value; }
            get { return _connectString; }
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
