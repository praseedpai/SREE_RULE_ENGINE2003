using System;
using System.Data;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;  

namespace ForeCastWatcher
{
	/// <summary>
	/// Updation to Oracle from the forecast involves lot of 
	/// steps. This class is written to make the updation transactional
	/// </summary>
	public class CBatchDBUpdator
	{
		/// <summary>
		///   Oracle Connection Object
		/// </summary>
		OracleConnection m_conn;
		/// <summary>
		///    For Transaction Processing
		/// </summary>
		OracleTransaction    m_Trans;
		/// <summary>
		///    Connection String
		/// </summary>
		String m_OrclConnstr ;
		/// <summary>
		///     Ctor
		///     Takes the connection string as parameter
		///     Opens a connection and starts transaction
		/// </summary>
		/// <param name="OrclConnstr"></param>
    	public CBatchDBUpdator(String OrclConnstr)
		{
			try 
			{
				m_OrclConnstr = OrclConnstr;
				m_conn = new OracleConnection(m_OrclConnstr);
				m_conn.Open();
				GlobalSettings.oracle_connection_count ++;
				m_Trans = m_conn.BeginTransaction(IsolationLevel.ReadCommitted); 
			}
			catch( Exception e )
			{
               System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString());  
               CSyntaxErrorLog.AddLine(e.ToString());
               throw e;
			}
		}

		
		/// <summary>
		///   Commit the transaction
		/// </summary>
	    public void Commit()
		{
			m_Trans.Commit(); 

		}
		/// <summary>
		///     Abort the Updation
		/// </summary>
		public void Abort()
	    {
		   m_Trans.Rollback(); 
	    }
		/// <summary>
		///     Close the Updation
		/// </summary>
		public void Close()
		{
			if ( m_Trans != null )
			{
				
				m_Trans.Dispose();	
				m_Trans = null;
			}

			if ( m_conn != null )
			{
				GlobalSettings.oracle_connection_count--;
				m_conn.Close(); 
				m_conn.Dispose();
				m_conn = null;
			}
 
		}
		/// <summary>
		///    Select a Scalar value
		/// </summary>
		/// <param name="cmd"></param>

		public DataSet SelectScalar(OracleCommand cmd )
		{
			try 
			{ 
				cmd.Connection = m_conn;
				cmd.CommandTimeout = 90;
				DataSet ds = new DataSet();
				OracleDataAdapter da = new OracleDataAdapter(cmd);
				da.Fill(ds, "Results");
				return ds;
			}
			catch(Exception e )
			{
				System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString());
				CSyntaxErrorLog.AddLine(e.ToString());
				Close();
				throw new CParserException(-100,e.ToString(),-1);
			}

		}


		/// <summary>
		///    Delete the Data
		/// </summary>
		/// <param name="cmd"></param>
		public void DeleteData(OracleCommand cmd )
		{
			try 
			{ 
				cmd.Connection = m_conn;
				cmd.CommandTimeout = 90;
				cmd.ExecuteNonQuery(); 
			}
			catch(Exception e )
			{
				System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString());
				CSyntaxErrorLog.AddLine(e.ToString());
  				Close();
				throw new CParserException(-100,e.ToString(),-1);
			}

		}
		/// <summary>
		/// Insert data
		/// </summary>
		/// <param name="cmd"></param>


		public void InsertData(OracleCommand cmd )
		{
			try 
			{ 
				cmd.Connection = m_conn;
				cmd.CommandTimeout = 90;
				cmd.ExecuteNonQuery(); 
			}
			catch(Exception e )
			{
				System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString());
				CSyntaxErrorLog.AddLine(e.ToString());
				Close();
				throw new CParserException(-100,e.ToString(),-1);
			}

		}



	}

}
