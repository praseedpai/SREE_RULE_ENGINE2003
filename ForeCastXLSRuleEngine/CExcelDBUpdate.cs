using System;
using System.Data;
using System.Data.Common;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
    
namespace ForeCastWatcher
{
	
		/// <summary>
		/// This class will update the forecast data given in
		/// excel to the DB.
		/// The class will be invoked after a thorough validation
		/// by running a script against it
		/// </summary>
		public class CExcelDBUpdate
		{
			

			/// <summary>
			///    Oracle Connection Str
			///    Given as parameter
			///    
			/// </summary>
			String m_OrclConnstr;

			/// <summary>
			///    A class to manage bulk updation of 
			///    data
			/// </summary>
			CBatchDBUpdator batch_upd=null;
			/// <summary>
			///   Excel Data reader helper
			/// </summary>
			CExcelReader    m_xlsreader = null;


			///
			///
			///
			///

			
            String DayLightAdj;  
			


			/// <summary>
			///   Ctor
			///   Passes Connection string to Excel Jet Driver
			///   and ODP.net Provider
			/// </summary>
			public CExcelDBUpdate(String OrclConnstr,CExcelReader reader)
			{
				m_xlsreader = reader;
				m_OrclConnstr =  OrclConnstr.Substring(0); 

			}
			/// <summary>
			///   Clean up
			/// </summary>

			public void Cleanup()
			{
				if ( m_xlsreader != null )
				{
					m_xlsreader.Close(); 
					m_xlsreader = null;
				}

				if ( batch_upd != null )
				{
                     batch_upd.Close();
					 batch_upd = null;

				}
                
			}

			private int GetTimeZoneOffset(string timename)
			{
#if false
				
				////////////////////////////////////////////
				///
				///  Oracle Command object for this rotuine
				///
				OracleCommand orcmd_select = null;
				String select_sql = "SELECT GMT_OFFSET_MINUTES FROM TIMEZONE WHERE ";
				select_sql += "UPPER(TRIM(GMT_NAME)) = UPPER(TRIM(:time_zone_name)) ";

				orcmd_select = new OracleCommand(select_sql);
				orcmd_select.Parameters.Add("time_zone_name",OracleDbType.Varchar2);

				orcmd_select.Parameters["time_zone_name"].Value  =  timename;

				DataSet ds = batch_upd.SelectScalar(orcmd_select);
 
				Time_Offset = Convert.ToInt16(ds.Tables[0].Rows[0]["GMT_OFFSET_MINUTES"]);
 
                return Time_Offset;			
#else
				////////////////////////////////////////////
				///
				///  Oracle Command object for this rotuine
				///
				OracleCommand orcmd_select = null;

				String select_sql = "SELECT TZ_OFFSET(TIMEZONE_LOCATION.TIMEZONE_LOCATION_NAME) PT_MINUTES_STR ";
				select_sql += " FROM  TIMEZONE_LOCATION  INNER JOIN TIMEZONE ";
                select_sql += " ON TIMEZONE_LOCATION.TIMEZONE_ID = TIMEZONE.TIMEZONE_ID ";
                select_sql += " AND upper(TRIM(TIMEZONE.GMT_NAME)) = UPPER(TRIM(:time_zone_name))";

				
				orcmd_select = new OracleCommand(select_sql);
				orcmd_select.Parameters.Add("time_zone_name",OracleDbType.Varchar2);

				orcmd_select.Parameters["time_zone_name"].Value  =  timename;

				DataSet ds = batch_upd.SelectScalar(orcmd_select);

				
				DayLightAdj = ds.Tables[0].Rows[0]["PT_MINUTES_STR"].ToString();

				int sLength = DayLightAdj.Length;
				DayLightAdj = DayLightAdj.Remove(sLength-1,1);
            	return 0;
                

				
#endif   


			}

			/// <summary>
			///    Update All the data from the spread sheet
			/// </summary>
			/// <returns></returns>
			/// 
			public bool UpdateAll()
			{
				String country;
				String segment;
				String currency;
				double ordc; // order cancel factor
                String forecast_name;
				try 
				{
					
					batch_upd = new CBatchDBUpdator(m_OrclConnstr); 
					country = m_xlsreader.GetCellValue("CONTROL","B2");
					currency  = m_xlsreader.GetCellValue("CONTROL","B5");
					segment = m_xlsreader.GetCellValue("CONTROL","B3");
					String ordcstr = m_xlsreader.GetCellValue("CONTROL","B6"); 
					string time_zone_code = m_xlsreader.GetCellValue("CONTROL","B4");
					GetTimeZoneOffset(time_zone_code);

                         
					ordcstr = ordcstr.Trim().ToUpper();
  
					int index = 0; 
					String temp ="";
					
					while ( char.IsDigit(ordcstr[index]) ) 
					{
						temp = temp + ordcstr[index];
						index++;
					}

					if ( ordcstr[index] == '.' )
					{
						index++;
						temp = temp + ".";
						while ( char.IsDigit(ordcstr[index]) ) 
						{
							temp = temp + ordcstr[index];
							index++;
						}

					}

					ordc = Convert.ToDouble(temp);
 	
					forecast_name = m_xlsreader.GetCellValue("CONTROL","B7"); 

					
					country = country.Trim().ToUpper();  
					segment = segment.Trim().ToUpper();
					currency = currency.Trim().ToUpper();
					forecast_name = forecast_name.Trim();
   
                    
					if ( country.Length == 0 || segment.Length == 0 
						|| currency.Length == 0 )
					{
						throw new Exception("Country , Segment or Currency is zero length"); 

					}

					if ( segment.Length > 10 )
						segment = segment.Substring(0,10);
 
					if (forecast_name.Length > 20 )
						forecast_name = forecast_name.Substring(0,20); 
									
					

                       
					upd_ob_revenue_units(country,segment,currency);
					update_click_stream(country,segment);
					update_ob_order_count(country,segment);
					update_odg_click_count(country,segment);
					update_odg_order_count_revenue(country,segment,currency);
					update_control_to_ei_cancel_factor(country,segment,forecast_name,ordc);
					batch_upd.Commit(); 
					batch_upd.Close();
					batch_upd = null;
					return true;
				}
				catch(CParserException e)
				{
					
					CSyntaxErrorLog.AddLine(e.ToString()); 
				    batch_upd.Abort();
                    batch_upd.Close();
					batch_upd = null;
					return false; 
				}
				catch(Exception e )
				{
					CSyntaxErrorLog.AddLine(e.ToString()); 
					batch_upd.Abort();
					batch_upd.Close();
			        batch_upd = null;
					return false; 

				}
			}
            /// <summary>
            ///    Find the number corresponding to Day of Week
            ///    0..6
            /// </summary>
            /// <param name="d"></param>
            /// <returns></returns>
 			public int DayOffset( System.DayOfWeek d )
			{
				switch(d)
				{
					case DayOfWeek.Sunday:
						return 0;
					case DayOfWeek.Monday:
						return 1;
					case DayOfWeek.Tuesday:
						return 2;
					case DayOfWeek.Wednesday:
						return 3;
					case DayOfWeek.Thursday:
						return 4;
					case DayOfWeek.Friday:
						return 5;
					case DayOfWeek.Saturday:
						return 6;
				}
				return -1;
			}
			/// <summary>
			///     Update OB Revenue units
			///     country ,segment and currency are the parameters
			/// </summary>
			/// <param name="Country"></param>
			/// <param name="segment"></param>
			/// <param name="Currency"></param>

			public void upd_ob_revenue_units(String country , 
				                             String segment,
				                             String currency )
			{
				
				////////////////////////////////////////////
				///
				///  Oracle Command object for this rotuine
				///

				OracleCommand orcmd_delete = null;
				OracleCommand orcmd_insert = null;

				//////////////////////////////////////////////////////////////////
				///
				///   SQL FOR DELETION OF OLD FORECAST DATA
				///
				///
				String delete_sql = "DELETE FROM OB_HOURLY_REVENUE_FORECAST WHERE ";
				delete_sql += "UPPER(COUNTRY_CODE) = UPPER(:country)  AND UPPER(EI_SEGMENT_CODE) =";
				delete_sql +="UPPER(:segment)   AND UDT_HOURLY_TIMESTAMP >= ";
				delete_sql +="TO_TIMESTAMP_TZ(TO_CHAR(:StartDate, 'DD-MON-YYHH24') || :DayLightAdj, 'DD-MON-RRHH24 TZH:TZM')";
				delete_sql +="AND  UDT_HOURLY_TIMESTAMP  < ";
				delete_sql +="TO_TIMESTAMP_TZ(TO_CHAR(:EndDate, 'DD-MON-YYHH24') || :DayLightAdj, 'DD-MON-RRHH24 TZH:TZM') AND ";
				delete_sql += "UPPER(CURRENCY_CODE) = UPPER(:currency) AND BRAND_ID = ";
				delete_sql += ":brandid";
				///////////////////////////////////////////////////////
				///
                ///  Stuff the Bind Variable 
                ///  

				orcmd_delete = new OracleCommand(delete_sql);
				orcmd_delete.Parameters.Add("country",OracleDbType.Varchar2);
				orcmd_delete.Parameters.Add("segment",OracleDbType.Varchar2);
				orcmd_delete.Parameters.Add("StartDate",OracleDbType.TimeStamp);  
				orcmd_delete.Parameters.Add("DayLightAdj",OracleDbType.Varchar2);   
				orcmd_delete.Parameters.Add("EndDate",OracleDbType.TimeStamp);    
				orcmd_delete.Parameters.Add("DayLightAdj",OracleDbType.Varchar2);   
				orcmd_delete.Parameters.Add("currency",OracleDbType.Varchar2);    
				orcmd_delete.Parameters.Add("brandid",OracleDbType.Int32);    

				//////////////////////////////////////////////
				///
				///   SQL for insertion
				///
				///
				String insert_sql= "INSERT into OB_HOURLY_REVENUE_FORECAST(country_code,EI_SEGMENT_CODE,";
				insert_sql+="brand_id,UDT_HOURLY_TIMESTAMP,currency_code,SUBTOTAL_REVENUE,TOTAL_QTY,UPDATE_DATE) values ";
				insert_sql+="(:country,:segment,:brandid,FROM_TZ(:udt_hourly_time_stamp,:DayLightAdj) ,:currency_code,:subtotal_revenue";
				insert_sql+=",:total_qty,:update_date)";


				
				String   [] COUNTRY     = new String[24];
				String   [] SEGMENT     = new String[24];
				int      [] BRANDID     = new int[24];
				DateTime [] DTTIME      = new DateTime[24];
				String   [] STR_LIGHT   = new String[24]; 
				String   [] CCODE       = new String[24];
				Decimal  [] SUB_TOTAL   = new Decimal[24];
				int      [] TOT_QTY     = new int[24];
				DateTime [] UPD_TIME    = new DateTime[24];
 
				orcmd_insert = new OracleCommand(insert_sql);
				orcmd_insert.ArrayBindCount = 24;

				OracleParameter prm = new OracleParameter("country", OracleDbType.Varchar2); 
				prm.Value     = COUNTRY;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("segment", OracleDbType.Varchar2); 
				prm.Value     = SEGMENT;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("brandid", OracleDbType.Int32); 
				prm.Value     = BRANDID;
				orcmd_insert.Parameters.Add(prm); 

                prm = new OracleParameter("udt_hourly_time_stamp", OracleDbType.TimeStamp); 
				prm.Value     = DTTIME;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("DayLightAdj", OracleDbType.Varchar2); 
				prm.Value     = STR_LIGHT;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("currency_code", OracleDbType.Char); 
				prm.Value     = CCODE;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("sub_total_revenue", OracleDbType.Decimal); 
				prm.Value     = SUB_TOTAL;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("total_qty", OracleDbType.Int32); 
				prm.Value     = TOT_QTY;
				orcmd_insert.Parameters.Add(prm);

				prm = new OracleParameter("upd_time", OracleDbType.Date); 
				prm.Value     = UPD_TIME;
				orcmd_insert.Parameters.Add(prm);
				

				

			
 				///////////////////////////////////////////////////////////
				///
				/// Find minimum and maximum date in OB revenue units
				///
				DateTime[] arr = m_xlsreader.GetMinMaxDate("OB_REVENUE_UNITS","DATE"); 
				///////////////////////////////////////////////////
				///
				/// Retrieve data from  OB revenue units
				///
				DataTable tab = m_xlsreader.GetDataTable("OB_REVENUE_UNITS");

				//////////////////////////////////////////////////
				///
				/// Iterate click stream worksheet
				///
				int tab_count = tab.Rows.Count;
				int iter=0;
				while ( iter < tab_count )
				{
					///////////////////////////////////
					///  Get the Current Row
					///
					DataRow rw = tab.Rows[iter];
					////////////////////////////////
					/// Retrieve brand lob and brand id
					///
					DateTime dt = Convert.ToDateTime(rw["DATE"]).Date;
					
 					int    brand_id = Convert.ToInt32(rw["BRAND_ID"]);
					orcmd_delete.Parameters["Country"].Value =  country;
					orcmd_delete.Parameters["Segment"].Value =  segment;
					orcmd_delete.Parameters["StartDate"].Value =dt;
					orcmd_delete.Parameters["DayLightAdj"].Value = DayLightAdj; 
					orcmd_delete.Parameters["EndDate"].Value =  dt.AddHours(24);
					orcmd_delete.Parameters["DayLightAdj"].Value = DayLightAdj; 
					orcmd_delete.Parameters["Currency"].Value = currency; 
					orcmd_delete.Parameters["brandid"].Value  = brand_id;
                    batch_upd.DeleteData(orcmd_delete);  
                    
					iter++;
				}





				int n_iter = 0; 	   	 
				int max_count = tab.Rows.Count;

				while ( n_iter < max_count )
				{

					DataRow rw = tab.Rows[n_iter];
					DateTime dt = Convert.ToDateTime(rw["DATE"]);

					String msf = dt.ToString(); 
					int column_offset = DayOffset( dt.DayOfWeek )*24;
					object [] double_data = 
						m_xlsreader.GetColumnValues("HOURLY_DISTRIBUTION",
						"Percent",column_offset,column_offset+23);
					double total_order = Convert.ToDouble(rw["REVENUE"]);
					int nbrand_id = Convert.ToInt32(rw["BRAND_ID"]);
 
                  
  
					String systemunits = rw["SYSTEM_UNITS"].ToString();
					double nsystem_units = 0; 
					if ( systemunits != "" )
					{
						nsystem_units = Convert.ToDouble(systemunits);  
					}

					int nhour = 0;

					while ( nhour < 24 )
					{
						COUNTRY[nhour]     = country;
						SEGMENT[nhour]     = segment;
						BRANDID[nhour]     = nbrand_id;
						dt = dt.AddHours((nhour == 0)?0:1);  
						DTTIME[nhour]      = dt;
						STR_LIGHT[nhour]  = DayLightAdj;
						CCODE[nhour]       = currency;
						SUB_TOTAL[nhour]   = Convert.ToDecimal(((total_order*Convert.ToDouble(double_data[nhour])))); 
						TOT_QTY[nhour]     = Convert.ToInt32(Math.Floor(((nsystem_units*Convert.ToDouble(double_data[nhour])))+0.5)); 
						UPD_TIME[nhour]    = DateTime.Now; 
						nhour++;
					}

					batch_upd.InsertData(orcmd_insert); 
					n_iter++; 

				}

				if ( orcmd_delete != null )
				{
                    orcmd_delete.Dispose();
					orcmd_delete = null;
				}

				if ( orcmd_insert != null )
				{
					orcmd_insert.Dispose();
					orcmd_insert = null;
				}
                      
				
			}

			/// <summary>
			///    Country and Segment 
			/// </summary>
			/// <param name="Country"></param>
			/// <param name="segment"></param>

			public void update_ob_order_count(String country , String segment )
			{

				////////////////////////////////////////////
				///
				///  Oracle Command object for this rotuine
				///
				OracleCommand orcmd_delete = null;
				OracleCommand orcmd_insert = null;

				///////////////////////////////////////////////////////
				///  Delete error string 
				///
				///
				String delete_sql = "DELETE FROM OB_HOURLY_COUNT_FORECAST WHERE ";
				delete_sql += "UPPER(COUNTRY_CODE) = UPPER(:country) AND UPPER(EI_SEGMENT_CODE) =";
				delete_sql +="UPPER(:segment)   AND UDT_HOURLY_TIMESTAMP >= ";
				delete_sql +="TO_TIMESTAMP_TZ(TO_CHAR(:StartDate, 'DD-MON-YYHH24') || :DayLightAdj, 'DD-MON-RRHH24 TZH:TZM')";
				delete_sql +="AND  UDT_HOURLY_TIMESTAMP  < ";
				delete_sql +="TO_TIMESTAMP_TZ(TO_CHAR(:EndDate, 'DD-MON-YYHH24') || :DayLightAdj, 'DD-MON-RRHH24 TZH:TZM') AND ";
				delete_sql += " UPPER(ONLINE_ORDER_TYPE) = UPPER(:pOnline) ";

				orcmd_delete = new OracleCommand(delete_sql);
				orcmd_delete.Parameters.Add("country",OracleDbType.Varchar2);
				orcmd_delete.Parameters.Add("segment",OracleDbType.Varchar2);
				orcmd_delete.Parameters.Add("StartDate",OracleDbType.TimeStamp); 
				orcmd_delete.Parameters.Add("DayLightAdj",OracleDbType.Varchar2);   
				orcmd_delete.Parameters.Add("EndDate",OracleDbType.TimeStamp);    
				orcmd_delete.Parameters.Add("DayLightAdj",OracleDbType.Varchar2);   
				orcmd_delete.Parameters.Add("pOnline",OracleDbType.Varchar2);   
								
				//////////////////////////////////////////////
				///
				///   SQL for insertion
				///
				///
				String insert_sql= "INSERT into OB_HOURLY_COUNT_FORECAST(country_code,EI_SEGMENT_CODE,";
				insert_sql+="UDT_HOURLY_TIMESTAMP,ONLINE_ORDER_TYPE,ORDER_COUNT,UPDATE_DATE) values ";
				insert_sql+="(:country,:segment,FROM_TZ(:udt_hourly_time_stamp,:DayLightAdj),:online_order_type,:order_count";
				insert_sql+=",:upd_date)";


				
				String   [] COUNTRY     = new String[24];
				String   [] SEGMENT     = new String[24];
				DateTime [] DTTIME      = new DateTime[24];
				String   [] STR_LIGHT   = new String[24]; 
				String   [] OOTYPE       = new String[24];
				int  [] ORD_COUNT   = new int[24];
				DateTime [] UPD_TIME    = new DateTime[24];
 
				orcmd_insert = new OracleCommand(insert_sql);
				orcmd_insert.ArrayBindCount = 24;

				OracleParameter prm = new OracleParameter("country", OracleDbType.Varchar2); 
				prm.Value     = COUNTRY;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("segment", OracleDbType.Varchar2); 
				prm.Value     = SEGMENT;
				orcmd_insert.Parameters.Add(prm); 

				
				prm = new OracleParameter("udt_hourly_time_stamp", OracleDbType.TimeStamp); 
				prm.Value     = DTTIME;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("DayLightAdj", OracleDbType.Varchar2); 
				prm.Value     = STR_LIGHT;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("online_order_type", OracleDbType.Varchar2); 
				prm.Value     = OOTYPE;
				orcmd_insert.Parameters.Add(prm); 

				
				prm = new OracleParameter("order_count", OracleDbType.Int32); 
				prm.Value     = ORD_COUNT;
				orcmd_insert.Parameters.Add(prm);

				prm = new OracleParameter("upd_time", OracleDbType.Date); 
				prm.Value     = UPD_TIME;
				orcmd_insert.Parameters.Add(prm);
				///////////////////////////////////////////////////////////
				///
				/// Find minimum and maximum date in ClickStream worksheet
				///

				DateTime[] arr = m_xlsreader.GetMinMaxDate("OB_ORDER_COUNT","DATE"); 

				


				
				///////////////////////////////////////////////////
				///
				/// Retrieve data from clickstream worksheet
				///
				DataTable tab = m_xlsreader.GetDataTable("OB_ORDER_COUNT");

				//////////////////////////////////////////////////
				///
				/// Iterate click stream worksheet
				///
				int tab_count = tab.Rows.Count;
				int iter=0;
				//int delete_count = 0;
				while ( iter < tab_count )
				{
					DataRow rw = tab.Rows[iter];
					DateTime dt = Convert.ToDateTime(rw["DATE"]).Date;
					
 
					String SOnline = Convert.ToString(rw["ONLINE_ORDER_TYPE"]);
					orcmd_delete.Parameters["country"].Value = country;
					orcmd_delete.Parameters["segment"].Value = segment;
					orcmd_delete.Parameters["StartDate"].Value = dt;
					orcmd_delete.Parameters["DayLightAdj"].Value = DayLightAdj; 
					orcmd_delete.Parameters["EndDate"].Value =   dt.AddHours(24);
					orcmd_delete.Parameters["DayLightAdj"].Value = DayLightAdj; 
					orcmd_delete.Parameters["pOnline"].Value = SOnline.ToUpper(); 
					batch_upd.DeleteData(orcmd_delete);  
					iter++;
				}
				
				int n_iter = 0; 	   	 
				int max_count = tab.Rows.Count;
				while ( n_iter < max_count )
				{
					DataRow rw = tab.Rows[n_iter];
					DateTime dt = Convert.ToDateTime(rw["DATE"]);
                    String SOnline = rw["ONLINE_ORDER_TYPE"].ToString();
					String msf = dt.ToString(); 
					int column_offset = DayOffset( dt.DayOfWeek )*24;
					object [] double_data = 
						m_xlsreader.GetColumnValues("HOURLY_DISTRIBUTION",
						"Percent",column_offset,column_offset+23);
					double total_order  =Convert.ToDouble(rw["ORDER_COUNT"]);
				
				

					int nhour = 0;
					while ( nhour < 24 )
					{
						COUNTRY[nhour]     = country;
						SEGMENT[nhour]     = segment;
						dt = dt.AddHours((nhour == 0)?0:1);  
						DTTIME[nhour]      = dt;
						STR_LIGHT[nhour] = DayLightAdj;
						OOTYPE[nhour]       = SOnline.ToUpper();
						ORD_COUNT[nhour]     = (int)Math.Floor(((total_order*Convert.ToDouble(double_data[nhour])))+0.5); 
						UPD_TIME[nhour]    = DateTime.Now; 
						nhour++;
					}

					batch_upd.InsertData(orcmd_insert); 
					n_iter++; 


				}

				if ( orcmd_delete != null )
				{
					orcmd_delete.Dispose();
					orcmd_delete = null;
				}

				if ( orcmd_insert != null )
				{
					orcmd_insert.Dispose();
					orcmd_insert = null;
				}
                       
			}
			/// <summary>
			///     odg_click_count
			/// </summary>
			/// <param name="Country"></param>
			/// <param name="segment"></param>

			public void update_odg_click_count(String country,String segment)
			{

				////////////////////////////////////////////
				///
				///  Oracle Command object for this rotuine
				///
				OracleCommand orcmd_delete = null;
				OracleCommand orcmd_insert = null;

				
				String delete_sql = "DELETE FROM ODG_HOURLY_COUNT_FORECAST WHERE ";
				delete_sql += "UPPER(COUNTRY_CODE) = UPPER(:country) AND EI_SEGMENT_CODE =";
				delete_sql +="UPPER(:segment)   AND UDT_HOURLY_TIMESTAMP >= ";
				delete_sql +="TO_TIMESTAMP_TZ(TO_CHAR(:StartDate, 'DD-MON-YYHH24') || :DayLightAdj, 'DD-MON-RRHH24 TZH:TZM')";
				delete_sql +="AND  UDT_HOURLY_TIMESTAMP  < ";
				delete_sql +="TO_TIMESTAMP_TZ(TO_CHAR(:EndDate, 'DD-MON-YYHH24') || :DayLightAdj, 'DD-MON-RRHH24 TZH:TZM') AND ";
				delete_sql +="UPPER(DGV_CODE) = UPPER(:dgv_code)";

				orcmd_delete = new OracleCommand(delete_sql);
				orcmd_delete.Parameters.Add("country",OracleDbType.Varchar2);
				orcmd_delete.Parameters.Add("segment",OracleDbType.Varchar2);
				orcmd_delete.Parameters.Add("StartDate",OracleDbType.TimeStamp);  
				orcmd_delete.Parameters.Add("DayLightAdj",OracleDbType.Varchar2);   
				orcmd_delete.Parameters.Add("EndDate",OracleDbType.TimeStamp);    
				orcmd_delete.Parameters.Add("DayLightAdj",OracleDbType.Varchar2);   
				orcmd_delete.Parameters.Add("dgv_code",OracleDbType.Varchar2);    

				//////////////////////////////////////////////
				///
				///   SQL for insertion
				///
				///
				String insert_sql= "INSERT into ODG_HOURLY_COUNT_FORECAST(country_code,EI_SEGMENT_CODE,";
				insert_sql+="DGV_CODE,UDT_HOURLY_TIMESTAMP,CLICK_COUNT,UPDATE_DATE) values ";
				insert_sql+="(:country,:segment,:dgv_code,FROM_TZ(:udt_hourly_time_stamp,:DayLightAdj),:click_count";
				insert_sql+=",:upd_date)";


				
				String   [] COUNTRY      =  new String[24];
				String   [] SEGMENT      =  new String[24];
                String   [] DGVCODE      =  new String[24];   
				DateTime [] DTTIME      = new DateTime[24];
				String   [] STR_LIGHT   = new String[24];  
				Decimal  [] CLICK_COUNT   = new Decimal[24];
				DateTime [] UPD_TIME    = new DateTime[24];

 
				orcmd_insert = new OracleCommand(insert_sql);
				orcmd_insert.ArrayBindCount = 24;
				OracleParameter prm = new OracleParameter("country", OracleDbType.Varchar2); 
				prm.Value     = COUNTRY;
				orcmd_insert.Parameters.Add(prm); 
				prm = new OracleParameter("segment", OracleDbType.Varchar2); 
				prm.Value     = SEGMENT;
				orcmd_insert.Parameters.Add(prm); 
			
				prm = new OracleParameter("dgv_code", OracleDbType.Varchar2); 
				prm.Value     = DGVCODE;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("udt_hourly_time_stamp", OracleDbType.TimeStamp); 
				prm.Value     = DTTIME;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("DayLightAdj", OracleDbType.Varchar2); 
				prm.Value     = STR_LIGHT;
				orcmd_insert.Parameters.Add(prm); 
							
				prm = new OracleParameter("click_count", OracleDbType.Int32); 
				prm.Value     = CLICK_COUNT;
				orcmd_insert.Parameters.Add(prm);

				prm = new OracleParameter("upd_time", OracleDbType.Date); 
				prm.Value     = UPD_TIME;
				orcmd_insert.Parameters.Add(prm);

			
				
				
				///////////////////////////////////////////////////////////
				///
				/// Find minimum and maximum date in ClickStream worksheet
				///
				DateTime[] arr = m_xlsreader.GetMinMaxDate("LINKTRACKER","DATE"); 

				


				



				
				///////////////////////////////////////////////////
				///
				/// Retrieve data from clickstream worksheet
				///
				DataTable tab = m_xlsreader.GetDataTable("LINKTRACKER");

				//////////////////////////////////////////////////
				///
				/// Iterate click stream worksheet
				///
				int tab_count = tab.Rows.Count;
				int iter=0;
				//int delete_count = 0;
				while ( iter < tab_count )
				{
					DataRow rw = tab.Rows[iter];
					DateTime dt = Convert.ToDateTime(rw["DATE"]).Date;
				
				    String dgv = Convert.ToString(rw["DGV_CODE"]); 
					orcmd_delete.Parameters["country"].Value   = country;
					orcmd_delete.Parameters["segment"].Value   = segment;
					orcmd_delete.Parameters["StartDate"].Value = dt;
					orcmd_delete.Parameters["DayLightAdj"].Value = DayLightAdj; 
					orcmd_delete.Parameters["EndDate"].Value   = dt.AddHours(24);
					orcmd_delete.Parameters["DayLightAdj"].Value = DayLightAdj; 
					orcmd_delete.Parameters["dgv_code"].Value  = dgv;
					batch_upd.DeleteData(orcmd_delete);  
					iter++;
				}



				
            
				///////////////////////////////////////////////////
				///
				///  Update spreadsheet data 
				///
				///
				
				int n_iter = 0; 	   	 
				int max_count = tab.Rows.Count;
				while ( n_iter < max_count )
				{
					DataRow rw = tab.Rows[n_iter];
					DateTime dt = Convert.ToDateTime(rw["DATE"]);
					String st = Convert.ToString(rw["ODG_CLICKS"]); 
                    String SOnline = Convert.ToString(rw["DGV_CODE"]);  

					String msf = dt.ToString(); 
					int column_offset = DayOffset( dt.DayOfWeek )*24;
					object [] double_data = 
						m_xlsreader.GetColumnValues("HOURLY_DISTRIBUTION",
						"Percent",column_offset,column_offset+23);
					double total_click_count  =Convert.ToDouble(rw["ODG_CLICKS"]);
					
                
					int nhour = 0;
					while ( nhour < 24 )
					{

						

						COUNTRY[nhour]     = country;
						SEGMENT[nhour]     = segment;
						DGVCODE[nhour]       = SOnline.ToUpper();
						dt = dt.AddHours((nhour == 0)?0:1);  
						DTTIME[nhour]      = dt;
						STR_LIGHT[nhour]   = DayLightAdj;
						CLICK_COUNT[nhour]     = Convert.ToDecimal(Math.Floor((total_click_count*Convert.ToDouble(double_data[nhour]))+0.5)); 
						UPD_TIME[nhour]    = DateTime.Now; 
						nhour++;
                         




					}
					batch_upd.InsertData(orcmd_insert); 
			
					n_iter++; 

				}

                       
				if ( orcmd_delete != null )
				{
					orcmd_delete.Dispose();
					orcmd_delete = null;
				}

				if ( orcmd_insert != null )
				{
					orcmd_insert.Dispose();
					orcmd_insert = null;
				}				 


			}
			/// <summary>
			///     
			/// </summary>
			/// <param name="country"></param>
			/// <param name="segment"></param>


			public void update_odg_order_count_revenue(String country,String segment,String currency)
			{

				////////////////////////////////////////////
				///
				///  Oracle Command object for this rotuine
				///
				OracleCommand orcmd_delete = null;
				OracleCommand orcmd_insert = null;
				/////////////////////////////////////////////////
				///
				///
				///
				///
				String delete_sql = "DELETE FROM ODG_HOURLY_REVENUE_FORECAST WHERE ";
				delete_sql += "UPPER(COUNTRY_CODE) = UPPER(:Country) AND EI_SEGMENT_CODE =";
				delete_sql +="UPPER(:segment)   AND UDT_HOURLY_TIMESTAMP >= ";
				delete_sql +="TO_TIMESTAMP_TZ(TO_CHAR(:StartDate, 'DD-MON-YYHH24') || :DayLightAdj, 'DD-MON-RRHH24 TZH:TZM')";
				delete_sql +="AND  UDT_HOURLY_TIMESTAMP  < ";
				delete_sql +="TO_TIMESTAMP_TZ(TO_CHAR(:EndDate, 'DD-MON-YYHH24') || :DayLightAdj, 'DD-MON-RRHH24 TZH:TZM') AND ";
				delete_sql +="UPPER(DGV_CODE) = UPPER(:sdgv_code) AND UPPER(CURRENCY_CODE) = UPPER(:Currency)";

				orcmd_delete = new OracleCommand(delete_sql);
				orcmd_delete.Parameters.Add("Country",OracleDbType.Varchar2);
				orcmd_delete.Parameters.Add("segment",OracleDbType.Varchar2);
				orcmd_delete.Parameters.Add("StartDate",OracleDbType.TimeStamp);  
				orcmd_delete.Parameters.Add("DayLightAdj",OracleDbType.Varchar2);   
				orcmd_delete.Parameters.Add("EndDate",OracleDbType.TimeStamp); 
				orcmd_delete.Parameters.Add("DayLightAdj",OracleDbType.Varchar2);   
				orcmd_delete.Parameters.Add("sdgv_code",OracleDbType.Varchar2);    
                orcmd_delete.Parameters.Add("Currency",OracleDbType.Varchar2);    

				

				
				//////////////////////////////////////////////
				///
				///   SQL for insertion
				///
				///
				String insert_sql= "INSERT into ODG_HOURLY_REVENUE_FORECAST(country_code,EI_SEGMENT_CODE,";
				insert_sql+="DGV_CODE,UDT_HOURLY_TIMESTAMP,ONLINE_ORDER_TYPE,CURRENCY_CODE,SUBTOTAL_REVENUE,ORDER_COUNT,UPDATE_DATE) values ";
				insert_sql+="(:country,:segment,:dgv_code,FROM_TZ(:udt_hourly_time_stamp,:DayLightAdj),:online_order_type,:currency,:subtotal_revenue,:order_count";
				insert_sql+=",:upd_date)";


				
				String   [] COUNTRY            =  new String[24];
				String   [] SEGMENT            =  new String[24];
				String   [] DGVCODE            =  new String[24];   
				DateTime [] DTTIME             =  new DateTime[24];
				String   [] STR_LIGHT          =  new String[24]; 
                String   [] OOTYPE             =  new String[24];       
				String   [] CURRENCY           =  new String[24]; 
				Decimal  [] SUB_TOTAL_REVENUE  =  new Decimal[24];
				Decimal  [] SUB_TOTAL_COUNT    =  new Decimal[24]; 
				DateTime [] UPD_TIME           =  new DateTime[24];

 
				orcmd_insert = new OracleCommand(insert_sql);
				orcmd_insert.ArrayBindCount = 24;
				OracleParameter prm = new OracleParameter("country", OracleDbType.Varchar2); 
				prm.Value     = COUNTRY;
				orcmd_insert.Parameters.Add(prm); 
				prm = new OracleParameter("segment", OracleDbType.Varchar2); 
				prm.Value     = SEGMENT;
				orcmd_insert.Parameters.Add(prm); 
			
				prm = new OracleParameter("dgv_code", OracleDbType.Varchar2); 
				prm.Value     = DGVCODE;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("udt_hourly_time_stamp", OracleDbType.TimeStamp); 
				prm.Value     = DTTIME;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("DayLightAdj", OracleDbType.Varchar2); 
				prm.Value     = STR_LIGHT;
				orcmd_insert.Parameters.Add(prm); 


				prm = new OracleParameter("online_order_type", OracleDbType.Varchar2); 
				prm.Value     = OOTYPE;
				orcmd_insert.Parameters.Add(prm); 

				prm = new OracleParameter("currency", OracleDbType.Varchar2); 
				prm.Value     = CURRENCY;
				orcmd_insert.Parameters.Add(prm); 
							
				prm = new OracleParameter("revenue_count", OracleDbType.Decimal); 
				prm.Value     = SUB_TOTAL_REVENUE;
				orcmd_insert.Parameters.Add(prm);

				prm = new OracleParameter("order_count", OracleDbType.Int32); 
				prm.Value     = SUB_TOTAL_COUNT;
				orcmd_insert.Parameters.Add(prm);


				prm = new OracleParameter("upd_time", OracleDbType.Date); 
				prm.Value     = UPD_TIME;
				orcmd_insert.Parameters.Add(prm);
				
				///////////////////////////////////////////////////////////
				///
				/// Find minimum and maximum date in ClickStream worksheet
				///
				DateTime[] arr = m_xlsreader.GetMinMaxDate("LINKTRACKER","DATE"); 

				


				

				


				
				///////////////////////////////////////////////////
				///
				/// Retrieve data from clickstream worksheet
				///
				DataTable tab = m_xlsreader.GetDataTable("LINKTRACKER");

				//////////////////////////////////////////////////
				///
				/// Iterate click stream worksheet
				///
				int tab_count = tab.Rows.Count;
				int iter=0;
				//int delete_count = 0;
				while ( iter < tab_count )
				{
					DataRow rw = tab.Rows[iter];
					DateTime dt = Convert.ToDateTime(rw["DATE"]).Date;
					String dgv = Convert.ToString(rw["DGV_CODE"]); 
                  

					
					

					orcmd_delete.Parameters["Country"].Value   = country;
					orcmd_delete.Parameters["Segment"].Value   = segment;
					orcmd_delete.Parameters["StartDate"].Value = dt;
                    orcmd_delete.Parameters["DayLightAdj"].Value = DayLightAdj; 
					orcmd_delete.Parameters["EndDate"].Value   = dt.AddHours(24);
					orcmd_delete.Parameters["DayLightAdj"].Value = DayLightAdj; 
					orcmd_delete.Parameters["sdgv_code"].Value  = dgv.ToUpper();
					orcmd_delete.Parameters["Currency"].Value  = currency;               
					batch_upd.DeleteData(orcmd_delete);  
					  
					          

                  
					iter++;
				}


				

				
            
				///////////////////////////////////////////////////
				///
				///  Update spreadsheet data 
				///
				///
				
				int n_iter = 0; 	   	 
				int max_count = tab.Rows.Count;
				while ( n_iter < max_count )
				{
					DataRow rw = tab.Rows[n_iter];
					DateTime dt = Convert.ToDateTime(rw["DATE"]);
					String st = Convert.ToString(rw["ODG_CLICKS"]); 
                    String SOnline = Convert.ToString(rw["DGV_CODE"]); 

					String msf = dt.ToString(); 
					int column_offset = DayOffset( dt.DayOfWeek )*24;
					object [] double_data = 
						m_xlsreader.GetColumnValues("HOURLY_DISTRIBUTION",
						"Percent",column_offset,column_offset+23);
					double total_order_count    =Convert.ToDouble(rw["ODG_ORDER_COUNT"]);
					double total_order_revenue  =Convert.ToDouble(rw["ODG_REVENUE"]);
					
                   
 
					int nhour = 0;
					while ( nhour < 24 )
					{
					
						COUNTRY[nhour]     = country;
						SEGMENT[nhour]     = segment;
						DGVCODE[nhour]       = SOnline.ToUpper();
						CURRENCY[nhour]    = currency;
                        		dt = dt.AddHours((nhour == 0)?0:1);  
						DTTIME[nhour]      = dt;
						STR_LIGHT[nhour]   = DayLightAdj;
						OOTYPE[nhour]      = "SYSTEM";
						SUB_TOTAL_COUNT[nhour]     = Convert.ToDecimal(Math.Floor(total_order_count*Convert.ToDouble(double_data[nhour])+0.5)); 
						SUB_TOTAL_REVENUE[nhour]   = Convert.ToDecimal(total_order_revenue*Convert.ToDouble(double_data[nhour])); 
						UPD_TIME[nhour]    = DateTime.Now; 
						nhour++;
                         
						
					}
					batch_upd.InsertData(orcmd_insert); 
					n_iter++; 

				}

                       
				if ( orcmd_delete != null )
				{
					orcmd_delete.Dispose();
					orcmd_delete = null;
				}

				if ( orcmd_insert != null )
				{
					orcmd_insert.Dispose();
					orcmd_insert = null;
				}	
				 


			}
			/// <summary>
			///    
			/// </summary>
			/// <param name="country"></param>
			/// <param name="segment"></param>
			/// <param name="ordc"></param>
			public void update_control_to_ei_cancel_factor(String country,String segment,
				String forecast_name , double ordc)
			{
                 
				
				
				////////////////////////////////////////////
				///
				///  Oracle Command object for this rotuine
				///
				OracleCommand orcmd_delete = null;
				OracleCommand orcmd_insert = null;

				String delete_sql = "DELETE FROM EI_CANCEL_FACTOR WHERE ";
				delete_sql += "UPPER(COUNTRY_CODE) = UPPER(:country)  AND UPPER(EI_SEGMENT_CODE) = ";
				delete_sql +=" UPPER(:segment) ";



				orcmd_delete = new OracleCommand(delete_sql);
				orcmd_delete.Parameters.Add("country",OracleDbType.Varchar2);
				orcmd_delete.Parameters.Add("segment",OracleDbType.Varchar2);
				orcmd_delete.Parameters["country"].Value = country;
				orcmd_delete.Parameters["segment"].Value = segment;
                batch_upd.DeleteData(orcmd_delete);  


				///////////////////////////////////////////////////
				///  Now insert new data
				///
				String insert_sql= "INSERT into EI_CANCEL_FACTOR(country_code,EI_SEGMENT_CODE,";
				insert_sql+="OB_CANCEL_FACTOR,FORECAST_NAME) values ";
				insert_sql+="(:country,:segment,:ob_cancel_factor";
				insert_sql+=",:forecast_name)";


				orcmd_insert = new OracleCommand(insert_sql);

				orcmd_insert.Parameters.Add("country",OracleDbType.Varchar2);
				orcmd_insert.Parameters.Add("segment",OracleDbType.Varchar2);
				orcmd_insert.Parameters.Add("ob_cancel_factor",OracleDbType.Decimal);
				orcmd_insert.Parameters.Add("forecast_name",OracleDbType.Varchar2);
 
				orcmd_insert.Parameters["country"].Value =  country; 
                orcmd_insert.Parameters["segment"].Value =  segment; 
                orcmd_insert.Parameters["ob_cancel_factor"].Value =  Convert.ToDecimal( Math.Round(ordc/100,3));
                orcmd_insert.Parameters["forecast_name"].Value =  forecast_name;
				batch_upd.InsertData(orcmd_insert); 

				if ( orcmd_delete != null )
				{
					orcmd_delete.Dispose();
					orcmd_delete = null;
				}

				if ( orcmd_insert != null )
				{
					orcmd_insert.Dispose();
					orcmd_insert = null;
				}

			}

			/// <summary>
			///   Update ClickStream Data
			/// </summary>
			public void update_click_stream(String country , String segment)
			{


				////////////////////////////////////////////
				///
				///  Oracle Command object for this rotuine
				///
				OracleCommand orcmd_delete = null;
				OracleCommand orcmd_insert = null;



				////////////////////////////////////////////////
				///  Delete SQL Code
				///
				///

				
				
				String delete_sql = "DELETE FROM CLICK_HOURLY_FORECAST WHERE ";
				delete_sql += "UPPER(COUNTRY_CODE) = UPPER(:country)  AND UPPER(EI_SEGMENT_CODE) = ";
				delete_sql +="UPPER(:segment)   AND UDT_HOURLY_TIMESTAMP >= ";
				delete_sql +="TO_TIMESTAMP_TZ(TO_CHAR(:StartDate, 'DD-MON-YYHH24') || :DayLightAdj, 'DD-MON-RRHH24 TZH:TZM')";
				delete_sql +="AND  UDT_HOURLY_TIMESTAMP  < ";
				delete_sql +="TO_TIMESTAMP_TZ(TO_CHAR(:EndDate, 'DD-MON-YYHH24') || :DayLightAdj, 'DD-MON-RRHH24 TZH:TZM') ";
				


				orcmd_delete = new OracleCommand(delete_sql);
				orcmd_delete.Parameters.Add("country",OracleDbType.Varchar2);
				orcmd_delete.Parameters.Add("segment",OracleDbType.Varchar2);
				orcmd_delete.Parameters.Add("StartDate",OracleDbType.TimeStamp);  
                orcmd_delete.Parameters.Add("DayLightAdj",OracleDbType.Varchar2);   
				orcmd_delete.Parameters.Add("EndDate",OracleDbType.TimeStamp);  
				orcmd_delete.Parameters.Add("DayLightAdj",OracleDbType.Varchar2);   
  
				

				//////////////////////////////////////////////
				///
				///   SQL for insertion
				///
				///
				String insert_sql= "INSERT into CLICK_HOURLY_FORECAST(country_code,EI_SEGMENT_CODE,";
				insert_sql+="UDT_HOURLY_TIMESTAMP,SITE_VISIT_COUNT,STORE_VISIT_COUNT,UPDATE_DATE) values ";
				insert_sql+="(:country,:segment,FROM_TZ(:udt_hourly_time_stamp,:DayLightAdj),:site_visit";
				insert_sql+=",:store_visit,:upd_date)";


				
				String   [] COUNTRY      =  new String[24];
				String   [] SEGMENT      =  new String[24];
				DateTime [] DTTIME       =  new DateTime[24];
				String   [] STR_LIGHT    =  new String[24];
				int      [] SITE_VISIT   =  new int[24];   
				int      [] STORE_VISIT  =  new int[24];
				DateTime [] UPD_TIME     =  new DateTime[24];

 
				orcmd_insert = new OracleCommand(insert_sql);
				orcmd_insert.ArrayBindCount = 24;
				OracleParameter prm = new OracleParameter("country", OracleDbType.Varchar2); 
				prm.Value     = COUNTRY;
				orcmd_insert.Parameters.Add(prm); 
				prm = new OracleParameter("segment", OracleDbType.Varchar2); 
				prm.Value     = SEGMENT;
				orcmd_insert.Parameters.Add(prm); 
				prm = new OracleParameter("udt_hourly_time_stamp", OracleDbType.TimeStamp); 
				prm.Value     = DTTIME;
				orcmd_insert.Parameters.Add(prm); 
				prm = new OracleParameter("DayLightAdj", OracleDbType.Varchar2); 
				prm.Value     = STR_LIGHT;
				orcmd_insert.Parameters.Add(prm); 
				prm = new OracleParameter("site_visit", OracleDbType.Int32); 
				prm.Value     = SITE_VISIT;
				orcmd_insert.Parameters.Add(prm); 
				prm = new OracleParameter("store_visit", OracleDbType.Int32); 
				prm.Value     = STORE_VISIT;
				orcmd_insert.Parameters.Add(prm);
				prm = new OracleParameter("upd_time", OracleDbType.Date); 
				prm.Value     = UPD_TIME;
				orcmd_insert.Parameters.Add(prm);

				
				///////////////////////////////////////////////////////////
				///
				/// Find minimum and maximum date in ClickStream worksheet
				///
				DateTime[] arr = m_xlsreader.GetMinMaxDate("CLICKSTREAM","DATE"); 




				
				///////////////////////////////////////////////////
				///
				/// Retrieve data from clickstream worksheet
				///
				DataTable tab = m_xlsreader.GetDataTable("CLICKSTREAM");

				//////////////////////////////////////////////////
				///
				/// Iterate click stream worksheet
				///
				int tab_count = tab.Rows.Count;
				int iter=0;
				//int delete_count = 0;
				while ( iter < tab_count )
				{
					DataRow rw = tab.Rows[iter];
					DateTime dt = Convert.ToDateTime(rw["DATE"]).Date;
					
					orcmd_delete.Parameters["Country"].Value = country;
					orcmd_delete.Parameters["Segment"].Value = segment;
					orcmd_delete.Parameters["StartDate"].Value = dt;
					orcmd_delete.Parameters["DayLightAdj"].Value = DayLightAdj; 
					orcmd_delete.Parameters["EndDate"].Value =   dt.AddHours(24);
					orcmd_delete.Parameters["DayLightAdj"].Value = DayLightAdj; 

					batch_upd.DeleteData(orcmd_delete);  
					iter++;
				}


				

            
				///////////////////////////////////////////////////
				///
				///  Update spreadsheet data 
				///
				///
				int n_iter = 0; 	   	 
				int max_count = tab.Rows.Count;
				while ( n_iter < max_count )
				{
					DataRow rw = tab.Rows[n_iter];
					DateTime dt = Convert.ToDateTime(rw["DATE"]);

					String msf = dt.ToString(); 
					int column_offset = DayOffset( dt.DayOfWeek )*24;
					object [] double_data = 
						m_xlsreader.GetColumnValues("HOURLY_DISTRIBUTION",
						"Percent",column_offset,column_offset+23);
					double total_site  =Convert.ToDouble(rw["TOTAL_SITE_VISITS"]);
					double total_store =Convert.ToDouble(rw["TOTAL_STORE_VISITS"]);
					
					

					int nhour = 0;
					while ( nhour < 24 )
					{
						

						COUNTRY[nhour]     = country;
						SEGMENT[nhour]     = segment;
						dt = dt.AddHours((nhour == 0)?0:1);  
						DTTIME[nhour]       = dt;
						STR_LIGHT[nhour]    = DayLightAdj; 
						SITE_VISIT[nhour]      = (int)Math.Floor(((total_site*Convert.ToDouble(double_data[nhour])))+0.5); 
						STORE_VISIT[nhour]     = (int)Math.Floor(((total_store*Convert.ToDouble(double_data[nhour])))+0.5); 
						UPD_TIME[nhour]    = DateTime.Now; 
						nhour++;
					}

			
					batch_upd.InsertData(orcmd_insert); 
					n_iter++; 

				}

				if ( orcmd_delete != null )
				{
					orcmd_delete.Dispose();
					orcmd_delete = null;
				}

				if ( orcmd_insert != null )
				{
					orcmd_insert.Dispose();
					orcmd_insert = null;
				}                        
				
				 

			}


		}

}
