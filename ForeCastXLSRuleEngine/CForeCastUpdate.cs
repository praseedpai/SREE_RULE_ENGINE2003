using System;
using System.IO;
using System.Collections; 
using System.Xml;
using System.Web;
using System.Web.Mail;

namespace ForeCastWatcher
{


	/// <summary>
	///    A class for sending mails to concerned people
	/// </summary>
	public class CNotification
	{
       /// <summary>
       ///     List of persons to send the email
       /// </summary>
       ArrayList m_emails = new ArrayList();
       /// <summary>
       ///   Subject success
       /// </summary>
       String m_Subject;
       /// <summary>
       ///    Subject Failure
       /// </summary>
       String m_SubjectF;
       /// <summary>
       ///    Body of the mail
       /// </summary>
       String m_Body; 
       /// <summary>
       ///   Server name or IP address
       /// </summary>
       String m_ServerName;
       /// <summary>
       ///    From address
       /// </summary>
       String m_from;

       /// <summary>
       /// To Address
       /// </summary>
 
       String m_to;
 
       /// <summary>
       ///    Ctor
       /// </summary>
	   public CNotification()
	   {
             

	   }
       /// <summary>
       ///     Add Email to the list
       /// </summary>
       /// <param name="id"></param>
       public void AddEmail(String id )
	   {
            m_emails.Add(id);  
		}
		/// <summary>
		///    From 
		/// </summary>
		public String From
		{
			set
			{
                this.m_from  = value;
			} 

			get
			{
				return m_from; 
			}


		}
		/// <summary>
		///   Server or IP address
		/// </summary>
		public String Server
		{
			set 
			{
                  m_ServerName = value;

			}

			get 
			{
                   return m_ServerName; 

			}

		}
		/// <summary>
		///    Subject
		/// </summary>

		public String SubjectSuccess
		{
			set 
			{
                m_Subject = value;
			}

			get 
			{
                 return m_Subject;  

			}

		}
		/// <summary>
		/// 
		/// </summary>

		public String SubjectFailure
		{

			set 
			{
				m_SubjectF = value;
			}

			get 
			{
				return m_SubjectF;  

			}


		}
		/// <summary>
		///    Body 
		/// </summary>
		/// 
		public String Body
		{
			set 
			{
				m_Body = value;
			}

			get 
			{
				return m_Body;  

			}

		}

		/// <summary>
		///    TO 
		/// </summary>
		/// 
		public String TO
		{
			set 
			{
				m_to = value;
			}

			get 
			{
				return m_to;  

			}

		}


		/// <summary>
		///    Dispatch the mail to the id given in 
		///    the Forecast excel spreadsheet
		/// </summary>
		public void SendToExcelEmail(int nStatus)
		{
			///////////////////////////////////
			///  if no mail is found return
			///
			if ( m_emails.Count == 0 )
			    	return;


			try 
			{   
				MailMessage mail = new MailMessage();
				mail.From = m_from;
				mail.Subject = (nStatus == 0 ) ?  m_Subject : m_SubjectF;
				mail.Priority = MailPriority.High;
				mail.BodyFormat = MailFormat.Text;
				mail.Body = m_Body;
				mail.To = TO; 
				SmtpMail.SmtpServer = m_ServerName;
				SmtpMail.Send(mail);
				return;
			}
			catch(Exception e )
			{
			    //////////////////////////////////////////////////
			    /// normally should not reach here 
			    /// may be due to an invalid id ...
				///
				///
                System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString()); 
				
                return;
			}


		}



		/// <summary>
		///    Dispatch the mail to the tech support dept.
		///    ids will be given in the Config.xml file
		/// </summary>
		public void Send(int nStatus)
		{
			////////////////////////////////////
			///
			///  if no id is given return
			///
			if ( m_emails.Count == 0 )
				    return;


			try 
			{ 
				MailMessage mail = new MailMessage();
				mail.From = m_from;
				mail.Subject = (nStatus == 0 ) ?  m_Subject : m_SubjectF;
				mail.Priority = MailPriority.High;
				mail.BodyFormat = MailFormat.Text;
				mail.Body = m_Body;

				////////////////////////////////////////
				///
				///  Concatenate the List of emails.
				///

				int count = m_emails.Count;
				String EmailList = "";
				for( int i=0; i	< count ; ++ i )
				{
					if ( i == count-1 )
						EmailList = EmailList + m_emails[i].ToString();
					else
						EmailList = EmailList +m_emails[i].ToString()+";"; 
				}
				mail.To = EmailList;
				///////////////////////////////////////////
				///
				///  Try to dispatch the mail
				///
				SmtpMail.SmtpServer = m_ServerName;
				SmtpMail.Send(mail);
				
			}
			catch(Exception  e)
			{
				//////////////////////////////////////////////////
				/// normally should not reach here 
				/// may be due to an invalid id ...
				///
				///
				System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString());
				return;
          
                 

			}

		}

	}
	/// <summary>
	///    A structure for storing details 
	///    regarding excel file
	/// </summary>
	public struct ForeCastFileInfo
	{
		/// <summary>
		/// Fully Qualified path name of the file
		///
		/// </summary>
		public String PathName;  
		/// <summary>
		///    .XLS File name
		/// </summary>
		public String XlsFile;  
	}
	/// <summary>
	///   This class will update the forecast
	/// </summary>
	public class CForeCastUpdate
	{
		/// <summary>
		///    Path where forecast data is found 
		/// </summary>
		private String  m_rootPath;
		/// <summary>
		///    Path to Xml file where rules are stored 
		/// </summary>
		private String  m_xmlfile;
		/// <summary>
		///   Instantiate the notifier
		/// </summary>
    	CNotification notify = new CNotification();
 		/// <summary>
		///    Ctor
		/// </summary>
		/// <param name="rootPath"></param>
		public CForeCastUpdate(String rulefile )
		{
			////////////////////////////////////////////
			/// force to make a copy
			///
			m_xmlfile = rulefile.Substring(0); 
			/////////////////////////////////
			///
			///  Read the Config.xml file
			///
			Initialize();
		}

        /// <summary>
        ///    Read the Stuff from the initialization file 
        /// </summary>
		private void Initialize()
		{
			try 
			{
				XmlNode scratch;
				XmlDocument doc = new System.Xml.XmlDocument();  
				doc.Load(m_xmlfile);  
				XmlNode config_nd = doc.SelectSingleNode("CONFIG"); 
				scratch = config_nd.SelectSingleNode("DECIMAL_ROUNDING");
				GlobalSettings.AddEntry("DECIMAL_ROUNDING",Convert.ToInt32(scratch.InnerText));
				scratch = config_nd.SelectSingleNode("ALLOW_EMAILS");
				GlobalSettings.AddEntry("ALLOW_EMAILS",scratch.InnerText);
				scratch = config_nd.SelectSingleNode("QUARANTINE_SHEETS");
				GlobalSettings.AddEntry("QUARANTINE_SHEETS",scratch.InnerText);
				scratch = config_nd.SelectSingleNode("BACKUP_SHEETS");
				GlobalSettings.AddEntry("BACKUP_SHEETS",scratch.InnerText);
				scratch = config_nd.SelectSingleNode("EVENT_LOGGING");
				GlobalSettings.AddEntry("EVENT_LOGGING",scratch.InnerText);


				//
				//  Retrieve notification parameters
				//
				//
				XmlNode notify_nd = config_nd.SelectSingleNode("NOTIFICATION");
				 scratch = notify_nd.SelectSingleNode("EMAIL_LIST");

				XmlNodeList iter = scratch.ChildNodes;
 
				foreach(XmlNode rnode in iter )
				{
					notify.AddEmail(rnode.InnerText);     
				}

				scratch = notify_nd.SelectSingleNode("SUBJECT_SUCCESS");
				notify.SubjectSuccess = scratch.InnerText;

				scratch = notify_nd.SelectSingleNode("SUBJECT_FAILURE");
				notify.SubjectFailure = scratch.InnerText;
  
				scratch = notify_nd.SelectSingleNode("SMTPSERVER");
				notify.Server = scratch.InnerText;

				scratch = notify_nd.SelectSingleNode("SENDER");
				notify.From  = scratch.InnerText;
				//
				//
				//  Retrieve oracle string
				//
				//
				XmlNode retr_node = config_nd.SelectSingleNode("ORACLE_CONNECTION_STRING");
				GlobalSettings.AddEntry("ORACLE_CONNECTION_STRING",retr_node.InnerText); 
				retr_node = config_nd.SelectSingleNode("SCRIPT");
				GlobalSettings.AddEntry("VALIDATION_SCRIPT",retr_node.InnerText); 
				retr_node = config_nd.SelectSingleNode("FORECAST_PATH");
				GlobalSettings.AddEntry("FORECAST_PATH",retr_node.InnerText ); 
				ProcessForeCastPath(retr_node);
				m_rootPath = retr_node.InnerText;  
				retr_node = config_nd.SelectSingleNode("EXCEL_CONNECTION_STRING"); 
				XmlNode temp_node = retr_node.SelectSingleNode("PROVIDER");
				GlobalSettings.AddEntry("EXCEL_PROVIDER",temp_node.InnerText );
				temp_node = retr_node.SelectSingleNode("PROPERTIES");
				GlobalSettings.AddEntry("EXCEL_PROPERTIES",temp_node.InnerText );

			}
			catch ( Exception e )
			{
                CSyntaxErrorLog.AddLine(e.ToString());
				throw e;

			}
		}


		public bool ProcessForeCastPath(XmlNode node)
		{
			try 
			{
				////////////////////////////////////////
				///  Retrieve the root directory 
				///
				String root_dir = node.InnerText;
 				if (!Directory.Exists(root_dir)) 
				{
					Directory.CreateDirectory(root_dir);
					
				}
				

#if false
				////////////////////////////////////////////////
				///
				///  Retrieve the Country names
				///
                XmlNode nd  = node.SelectSingleNode("COUNTRY_LIST");
                XmlNodeList nd_list = nd.SelectNodes("COUNTRY");
  
				int i=0;
				while ( i < nd_list.Count )
				{
					//////////////////////////////
					/// Retrieve Country node
					///
					XmlNode country_nd = nd_list.Item(i);

					/////////////////////////////////////
					///  Retrieve country name
					///  
                    String country_name = country_nd.SelectSingleNode("COUNTRY_NAME").InnerText;

					///////////////////////////////////////////////
					///  Try to create Folder for the country
					///
					///
  					if (!Directory.Exists(root_dir+"\\"+country_name) )
						Directory.CreateDirectory(root_dir+"\\"+country_name);

					XmlNode  Temp  = country_nd.SelectSingleNode("SEGMENT_LIST");

					XmlNodeList Temp_List = null;
					if  ( Temp != null ) 
					     Temp_List = Temp.ChildNodes;
 
					
					if (Temp_List != null  ) 
					{
 
						if ( Temp_List.Count > 0 )
						{
							int j=0;
							while (j < Temp_List.Count )
							{
								String segment = Temp_List.Item(j).InnerText;
								if (!Directory.Exists(root_dir+"\\"+country_name+"\\"+segment) )
									Directory.CreateDirectory(root_dir+"\\"+country_name+"\\"+segment);
								j++;
							}

 
						}
					}

					i++;
				}

#endif


			}
			catch(Exception e )
			{
                CSyntaxErrorLog.AddLine(e.ToString());
				throw e;
			}

            return true;

		} 
		/// <summary>
		///    Date as Str . this routine is written to create
		///    a unique file name 
		/// </summary>
		/// <param name="dt"></param>
		/// <returns></returns>

		private String DateAsStr( DateTime dt )
		{
			return dt.Year.ToString() +"_"+
				dt.Month.ToString()+  "_"+
				dt.Day.ToString() +   "_"+  
				dt.Hour.ToString() +  "_"+
				dt.Minute.ToString()+ "_"+
				dt.Second.ToString()+ "_"+
				dt.Millisecond.ToString(); 

		}


		////////////////////////////////////////
		///
		///
		///
		public String GetCountryFromPath(String Path)
		{
			char[] splitter = { '\\' };

			String[] Temp = Path.Split( splitter,10);
            int nlen = Temp.Length;
			return Temp[nlen-3].Trim().ToUpper(); 

            
	   }

		///////////////////////////////////////////////
		///
		///
		///
		public String GetSegmentFromPath(String Path )
		{
			char[] splitter = { '\\' };
			String[] Temp = Path.Split( splitter,10);
			int nlen = Temp.Length;
			return Temp[nlen-2].Trim().ToUpper(); 

		}
        /// <summary>
		///    Call  the Excel content Validator
		/// </summary>
		/// <param name="s"></param>
		public bool ValidateFile(ForeCastFileInfo s)
		{

			CParser prs = null;
			CExcelDBUpdate upd=null;

			try {

			   if (!ForeCastUpdatorService.CanFilebeAccessed(s.XlsFile) )
			   {
				  throw new Exception(s.XlsFile + " Cannot be accessed or opened by another app");    
			   }

			   String m_Country = this.GetCountryFromPath(s.XlsFile);
               String m_Segment = this.GetSegmentFromPath(s.XlsFile);
 

               System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Processing File ...."+ s.XlsFile);   
			   String Prgs = (String)GlobalSettings.GetEntry("VALIDATION_SCRIPT");
			   String OrclConStr = (String)GlobalSettings.GetEntry("ORACLE_CONNECTION_STRING");
			   String ExcelConStr = "Provider="+(String)GlobalSettings.GetEntry("EXCEL_PROVIDER");
			   ExcelConStr += "Data Source="+ s.XlsFile+";";
			   ExcelConStr += (String)GlobalSettings.GetEntry("EXCEL_PROPERTIES");
               CSemanticErrorLog.AddLine("Excel file = " + s.XlsFile );
  			   prs = new CParser(Prgs,m_Country,m_Segment);    
			   bool brs = prs.CallParser(ExcelConStr,OrclConStr);
			   bool ret_val = true;
 
				if ( brs == true )
				{

					upd = new CExcelDBUpdate(OrclConStr,prs.GetExcelReader());  
					ret_val = upd.UpdateAll();
                     
					if ( ret_val == false )
					{
                        
						CSemanticErrorLog.AddLine("FATAL ERROR WHILE UPDATING FORECAST DATA  ");

					}


				}



				if ( ret_val ==  true  && brs == true )
				{
			
					if ( GlobalSettings.allow_emails == true )
					{
						notify.TO = prs.GetEmail();
						notify.Body = "Hi" +"\r\n\r\n" + "ForeCast updation status " + "\r\n"+prs.GetSemanticLog();  
						notify.SendToExcelEmail(0);
					}

					if ( GlobalSettings.event_logging == true )
					{
						if(!System.Diagnostics.EventLog.SourceExists("ForeCastLog"))
							System.Diagnostics.EventLog.CreateEventSource("ForeCastLog",
								"DoForeCastLog");

						System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Hi" +"\r\n\r\n" + "ForeCast updation status " + "\r\n"+prs.GetSemanticLog() + prs.GetSyntaxLog());  


					}

					if ( upd != null )
					{
                          upd.Cleanup();
						  upd = null;

					}
					if (prs != null )  
					{
						prs.CloseAll();  
						prs = null;
					}
					return ret_val;
				
				}
				else 
				{

					if ( GlobalSettings.event_logging == true )
					{
						if(!System.Diagnostics.EventLog.SourceExists("ForeCastLog"))
							System.Diagnostics.EventLog.CreateEventSource("ForeCastLog",
								"DoForeCastLog");

						System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Hi" +"\r\n\r\n" + "ForeCast updation status " + "\r\n"+prs.GetSemanticLog() + prs.GetSyntaxLog());  

					}

					if ( GlobalSettings.allow_emails == true )
					{
				
						notify.TO = prs.GetEmail();
						notify.Body = "Hi" +"\r\n\r\n" + "ForeCast updation status " + "\r\n"+prs.GetSemanticLog();  
						notify.SendToExcelEmail(1);

						notify.Body = "Hi" +"\r\n\r\n" + "ForeCast updation status " + "\r\n"+prs.GetSemanticLog() + prs.GetSyntaxLog();  
						notify.Send(1);
					}
					if ( upd != null )
					{
						upd.Cleanup();
						upd = null;

					}
					if (prs != null )  
					{
						prs.CloseAll();  
						prs = null;
					}

					return false;

				}
			}
			catch(Exception e )
			{
                System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString()); 
				CSemanticErrorLog.AddLine(e.ToString()); 
				if ( GlobalSettings.event_logging == true )
				{
					if(!System.Diagnostics.EventLog.SourceExists("ForeCastLog"))
						System.Diagnostics.EventLog.CreateEventSource("ForeCastLog",
							"DoForeCastLog");


					System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Hi" +"\r\n\r\n" + "ForeCast updation status " + "\r\n"+CSemanticErrorLog.GetLog() + CSyntaxErrorLog.GetLog());  

				}

				if ( GlobalSettings.allow_emails == true )
				{
				
					
					notify.Body = "Hi" +"\r\n\r\n" + "ForeCast updation status " + "\r\n"+CSemanticErrorLog.GetLog() + CSyntaxErrorLog.GetLog();  
					notify.Send(1);
				}

				if ( upd != null )
				{
					upd.Cleanup();
					upd = null;

				}
				if (prs != null )  
				{
					prs.CloseAll();  
					prs = null;
				}

				return false;

			}

           
			

		}

		
		/// <summary>
		///    Backup File
		/// </summary>
		/// <param name="s"></param>

		public void BackupFile(ForeCastFileInfo s )
		{

			try 
			{
			  String DirName = s.PathName+"\\"+"BACKUP\\";
			  String XlFile  = s.XlsFile;  
			  int nindex = XlFile.LastIndexOf("\\");
			  String filename = XlFile.Substring(nindex+1);
			  String newfilename = DirName + 
			                       DateAsStr(DateTime.Now)+"_"+
				                   filename;
              /////////////////////////////////////////////
              ///
			  ///  IF Backup folder does not exist
			  ///  Create a new one 
			  if (!Directory.Exists(DirName))
			  {
					Directory.CreateDirectory(DirName);
			  }
              //////////////////////////////////////
              ///
			  ///  Move to Backup folder
			  ///
			  ///

              System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Copying File"); 
			  if (!ForeCastUpdatorService.CanFilebeAccessed(XlFile) )
			  {
					throw new Exception(XlFile + "Copy: Cannot be accessed or opened by another app");    
			  }
              File.Copy(XlFile,newfilename,true);
			  System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Finished Copying File");  
			  if (!ForeCastUpdatorService.CanFilebeAccessed(XlFile) )
			  {
					throw new Exception(XlFile + "Delete: Cannot be accessed or opened by another app");    
			  }
			  File.Delete(XlFile);
              System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Finished Deleting File");     

			}
			catch ( Exception  e)
			{
              System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString());      
              CSyntaxErrorLog.AddLine(e.ToString());
              return;

			}
         
		}
		
		/// <summary>
		///   Move the spread sheet File to Quarantine
		/// </summary>
		/// <param name="s"></param>
		public void MoveToQuarantine(ForeCastFileInfo s )
		{


			try 
			{
			  String DirName = s.PathName+"\\"+"QUARANTINE\\";
			  String XlFile  = s.XlsFile;  
			  int nindex = XlFile.LastIndexOf("\\");
			  String filename = XlFile.Substring(nindex+1);
			  String newfilename = DirName + 
			                       DateAsStr(DateTime.Now)+"_"+
			                       filename;
              /////////////////////////////////////////////
              ///  
			  ///   IF Quarantine Folder does not exist 
			  ///   Create a new one
			  if (!Directory.Exists(DirName))
			   {
			    	Directory.CreateDirectory(DirName);
			   }
			  ///////////////////////////////////////////
			  ///
			  ///  Copy the File to Quarantine
			  ///
			  if (!ForeCastUpdatorService.CanFilebeAccessed(XlFile) )
			  {
					throw new Exception(XlFile + " Copy: Cannot be accessed or opened by another app");    
			  }
 			  File.Copy(XlFile,newfilename,true);
              //////////////////////////////////////////////
              ///  Delete the file 
			  ///
			  if (!ForeCastUpdatorService.CanFilebeAccessed(XlFile) )
				{
					throw new Exception(XlFile + " Cannot be accessed or opened by another app");    
				} 
			  File.Delete(XlFile);
			}
			catch(Exception  e)
			{
				///////////////////////////////////////////////
				///
				/// Log the Exception
				///
				System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString());
				CSyntaxErrorLog.AddLine(e.ToString());
  				return;  

			}
		}

		/// <summary>
		///    Update Routine
		///    Loop as long as a valid file is found
		/// </summary>
		public bool Update()
		{
           
			bool bVal = true;
		    while ( true )
		    {
				CDirectoryWalker x = new CDirectoryWalker();
				x.Visit(m_rootPath,"*.XLS",true);
				ArrayList r = x.GetAllXls(); 

                ///////////////////////////////
                ///  if no more file is found ...
                ///  stop this iteration
				///
				if ( r.Count == 0 )
					    break; 

				foreach( ForeCastFileInfo s in r )
				{      

					////////////////////////////////////////////////////
					///
					///  Stop processing , if thread has been requested to stop processing 
					///
					///
					if ( GlobalSettings.request_processing_thread_abort == true )
						            return bVal;
					///////////////////////////////
					/// Clean up the Error Logs before
					/// we go for fresh processing 
					///	CSemanticErrorLog.Cleanup();
					CSyntaxErrorLog.Cleanup();
					CSemanticErrorLog.Cleanup(); 
 
					////////////////////////////////////////
					///
					///  Try to Validate the file
					///
					GlobalSettings.oracle_connection_count = 0;
					GlobalSettings.excel_connection_count = 0;
					if ( ValidateFile(s) ) 
					{
                        System.GC.Collect();
  
						///////////////////////////////////////////////
						///
						///  Validation and updation is successful
						///  Check for the backup_sheets flag and 
						///  Back up the File , if it is true
						///  

						if ( GlobalSettings.oracle_connection_count != 0 )
						{
							System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Oracle Connection"+Convert.ToString(GlobalSettings.oracle_connection_count) );
						}


						if ( GlobalSettings.excel_connection_count != 0 )
						{
                           System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Oracle Connection"+Convert.ToString(GlobalSettings.excel_connection_count) );

						}
							 

						if ( GlobalSettings.backup_sheets == true )
							           BackupFile(s);
						
					}
					else 
					{

						///////////////////////////////////
						/// At least one stuff failed in this iteration
						/// left unused
						///
						if (bVal)
							bVal = false;

						///////////////////////////////////////////////////
						///
						///  Failure in updating ForeCast Data
						///  Try to Quarantine the file , if 
						///  quarantine_sheets flag is true 
						///  

						if ( GlobalSettings.oracle_connection_count > 0 )
						{
							System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Oracle Connection"+Convert.ToString(GlobalSettings.oracle_connection_count) );
						}


						if ( GlobalSettings.excel_connection_count > 0 )
						{
							System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Oracle Connection"+Convert.ToString(GlobalSettings.excel_connection_count) );

						}
						if ( GlobalSettings.quarantine_sheets == true )
                       		     MoveToQuarantine(s);    
						    
					}

					   
				}

			  break;
 
			}  

			return bVal;

		}

	}
}
