using System;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using Microsoft.Win32;

namespace Newsoft.DBConn
{
	/// <summary>
	/// ����������ͨ��OleDb����Oracle���ݿ⣬ʵ�ֲ�ѯ��������ݿ����
	/// �������ڣ�2005��11��
	/// �汾���У���Ӫ������Ұ������޹�˾
	/// </summary>
	public class ConnOracle //: IDisposable

	{
		// ���ݿ������ַ���
		private string _connectString		= "" ;	

        // ������Ϣ	
		private string _errorMessage		= "" ;		
	
		/// <summary>
		/// �������ݿ�
		/// </summary>
		public ConnOracle()
		{
			//
			// TODO: �ڴ˴���ӹ��캯���߼�
			//	
	
			//ȡ�������ַ���OracleConnectionString
            _connectString = System.Configuration.ConfigurationManager.AppSettings["OracleConnectionString"];

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

        public void Dispose()
        {
            GC.SuppressFinalize(true);
        }

		
		/// <summary>
        /// ���������ַ����е���Ϣ
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
        /// ��ȡ������Ϣ 
        /// </summary>
        public string ErrorMessage
        {
            get
            {
                return this._errorMessage;
            }
        }

        /// <summary>
		/// ִ�����ݿ��ѯ����
		/// </summary>
		/// <param name="_strSql"></param>
		/// <returns></returns>
		public  DataSet executeQuery(string _strSql)
		{
			//��������
			DataSet ds = new DataSet();	

            //OLEDB��������
            OleDbConnection Conn= new OleDbConnection(_connectString);
            OleDbDataAdapter objAdp= new OleDbDataAdapter(_strSql,Conn);
	
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

                if (Conn.State == ConnectionState.Open)
                {
                    Conn.Close();
                    Conn.Dispose();
                }
			}
			return ds;		
		}
        
		/// <summary>
		/// ִ�����ݿ���²���
		/// </summary>
		/// <param name="_strSql">SQL���</param>
		public void executeUpdate(string _strSql)
		{			
           //�������ݿ�����
            OleDbConnection Conn = new OleDbConnection(_connectString);		

            // �����ݿ�����
            Conn.Open();

            //�������ݿ��������
            OleDbCommand objCmd = new OleDbCommand(_strSql,Conn);

			try
			{
				//ִ�����ݿ���²���
				objCmd.ExecuteNonQuery();
			}
			catch(Exception ex)
			{
                _errorMessage = ex.Message;
                throw new Exception("���ݿ���²����쳣��" + ex.Message);
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
        /// ִ�����ݿ��ѯ����
        /// </summary>
        /// <param name="_strSql">��ѯ���</param>
        /// <returns></returns>
        public byte[] GetBlob(string _strSql)
        {
            // ���Blob����
            byte[] blob = new byte[0];

            //OLEDB��������
            OleDbConnection Conn = new OleDbConnection(_connectString);
            OleDbDataAdapter objAdp = new OleDbDataAdapter(_strSql, Conn);

            try
            {
                //��������
                DataSet ds = new DataSet();

                //��������		
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
                throw new Exception("���ݿ��ȡBLOB�ֶβ����쳣��" + ex.Message);
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
