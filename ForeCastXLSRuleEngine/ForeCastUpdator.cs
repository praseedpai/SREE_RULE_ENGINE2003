using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.IO;
using System.Threading; 

namespace ForeCastWatcher
{
	/// <summary>
	///    A Windows service to monitor Business forecast data on a 
	///    folder having specified structure. 
	/// </summary>
	public class ForeCastUpdatorService : System.ServiceProcess.ServiceBase
	{
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		/// <summary>
		///    Directory Path where service is running 
		/// </summary>

		private static String DirPath ;


		////////////////////////////////////////////////////
		///  
		///  Processing Logic Thread
		///

		private static Thread m_process;

		/// <summary>
		///     Ctor
		/// </summary>
		/// 
		public ForeCastUpdatorService()
		{
			// This call is required by the Windows.Forms Component Designer.
			InitializeComponent();

			
			////////////////////////////////////////////////
			///
			///  Try to create Event Log
			///
			if(!System.Diagnostics.EventLog.SourceExists("ForeCastLog"))
				System.Diagnostics.EventLog.CreateEventSource("ForeCastLog",
					"DoForeCastLog");

						
		}

		/// <summary>
		///    Entry Point for the Service
		/// </summary>
		static void Main()
		{
			System.ServiceProcess.ServiceBase[] ServicesToRun;
	
			

#if false

			/////////////////////////////////////////////////
			///
			///  Stub to Test the business logic 
			///
			///
			///try 
			///{
			///	CForeCastUpdate fcst_up = new CForeCastUpdate("C:\\ForeCastWatcher\\ForeCastWatcher\\CONFIG.XML");
			///	fcst_up.Update();
			///}
			///catch(Exception e )
			///{
            ///      e.ToString(); 
            /// 
		    ///}
		   	///return;

#endif

			ServicesToRun = new System.ServiceProcess.ServiceBase[] { new ForeCastUpdatorService() };
			System.ServiceProcess.ServiceBase.Run(ServicesToRun);



		}

		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			// 
			// ForeCastUpdatorService
			// 
			this.ServiceName = "ForeCastValidatorService";

		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			
			base.Dispose( disposing );
		}

		/// <summary>
		///    Scan for Files and Process if at all at least one is found
		/// </summary>
		/// 
		private static void ScanAndProcess()
		{

			CForeCastUpdate fcst_up = null;
			try 
			{
                
				System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Scanning for data");

				if (!ForeCastUpdatorService.CanFilebeAccessed(DirPath+"\\CONFIG.XML")) 
				{
                   String file_name = DirPath+"\\CONFIG.XML";
                   System.Diagnostics.EventLog.WriteEntry("ForeCastLog",file_name+"  does  not exist ,Cannot be accessed or opened already by another app");
                   return;
				}
					                  

				fcst_up = new CForeCastUpdate(DirPath+ "\\CONFIG.XML");

				try 
				{
					System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Going to iterate the folders");
					fcst_up.Update();
					System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Finished Scanning for data");
				}
				catch(Exception ex )
				{
                  
					System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Exception thrown"+ex.ToString()) ;   

				}

			}
			catch(Exception e )
			{
	      
				  System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Exception thrown"+e.ToString()) ;
				  
			
			}
			
			fcst_up = null;
			return;

		}

		////////////////////////////////////////////////////////
		///
		///  Here is the actual workhorse 
		///
		///
		public static void ScanForForeCastData() 
		{
			try
		    {
				while ( true )
				{
					    ////////////////////////////////
					    ///
					    ///  Clean up all the data
					    ///  
					    CSyntaxErrorLog.Cleanup();
					    CSemanticErrorLog.Cleanup();
					    GlobalSettings.Cleanup();
					    ///////////////////
					    ///
					    ///  Do Processing 
					    ///  

						ScanAndProcess();
					    /////////////////////////////////////////
					    ///
					    ///  Sleep the Thread for 2 minutes.
					    ///  in the production build , it should run
					    ///  in an interval of at least 10 minutes
					    ///
					    ///
					    System.GC.Collect();   
						System.Threading.Thread.Sleep(60000*5);     
					}
		
			 
		   }
		   catch(Exception e)
		   {

				System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Exception thrown by the Thread");
			    System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString());
		   }

           return;   


		}


		/// <summary>
		/// 
		/// </summary>
		/// <param name="pathname"></param>
		/// <returns></returns>

		public static bool CanFilebeAccessed(String pathname )
		{
			if (!File.Exists(pathname) )
			{
				
				return false;
			}
      			bool rtnvalue = true;
	        	try
	    		{
					System.IO.FileStream fs = System.IO.File.OpenWrite(pathname);
				    fs.Close();
    			}
	    		catch(System.IO.IOException )
		    	 {
				    rtnvalue = false;
			     }
    			return(rtnvalue);
				    

		}
		

		/// <summary>
		/// Set things in motion so your service can do its work.
		/// </summary>
		protected override void OnStart(string[] args)
		{
			
			////////////////////////////////////////////////////////////
			///  Retrieve the Path where service is being installed
			///
			System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Going to Abort the Processing ThreaD");   
			DirPath = System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
			m_process  = new Thread(new ThreadStart(ScanForForeCastData));
			m_process.Start(); 
					
		}
 
		/// <summary>
		/// Stop this service.
		/// </summary>
		protected override void OnStop()
		{
			
			try 
			{
				System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Going to Abort the Processing ThreaD");   
				m_process.Abort(); 
				GlobalSettings.request_processing_thread_abort = true;
				System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Finished Sending Abort Data");  
				System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Going to Call Join");
				m_process.Join();
                System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Processing Thread correctly terminated");			
				System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Excel forecast service stopped"); 		
			}

			catch(Exception e )
			{
				// Send a mail that service has been stopped
                System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString());
                System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Abort or Join Failed");  
			}

		    

		
			
		}

		
	}
}
