using System;
using System.IO;
using System.Collections;

namespace ForeCastWatcher
{
	/// <summary>
	/// Summary description for CDirectoryWalker.
	/// </summary>
	public class CDirectoryWalker
	{
		/// <summary>
		///    Array List to Store the number of xls file
		/// </summary>
		ArrayList xls = new ArrayList();
		/// <summary>
		///    Ctor
		/// </summary>
		public CDirectoryWalker()
		{
			
		}
        /// <summary>
        ///    Retrieve all the Xls . Call this routine
        ///    after Visit routine
        /// </summary>
        /// <returns></returns>
		public ArrayList GetAllXls()
		{
		   return xls;
 		}
		/// <summary>
		///    Reset 
		/// </summary>

		public void Reset() 
		{
			xls.Clear();
		}

		/// <summary>
		///    Visit all the directories and return a list of files
		///    to process
		/// </summary>
		/// <param name="Path"></param>
		/// <param name="Filter"></param>
		/// <param name="first_time"></param>
		public void  Visit(String Path , String Filter , bool first_time)
		{
			
			DirectoryInfo di = new DirectoryInfo(Path);
			
			DirectoryInfo[] dirs = di.GetDirectories("*.*");
			

			if ( first_time )
			{
				System.IO.FileInfo[] nts =  di.GetFiles(Filter);
				foreach(FileInfo ns in nts )
				{
					ForeCastFileInfo fi = new ForeCastFileInfo();
 					fi.XlsFile = ns.FullName;
					fi.PathName = ns.DirectoryName;
 					xls.Add(fi);  
				}
				first_time = false;
			}

            // Iterate through all the directories
			//
			//
			foreach (DirectoryInfo diNext in dirs) 
			{

				if ( diNext.Name.ToUpper()  == "BACKUP"  || 
					 diNext.Name.ToUpper()  == "QUARANTINE"  )
					           continue;
				
				System.IO.FileInfo[] nts =  diNext.GetFiles(Filter);
				
				foreach(FileInfo ns in nts )
				{

					ForeCastFileInfo fi = new ForeCastFileInfo();
					fi.XlsFile = ns.FullName;
					fi.PathName = ns.DirectoryName;
					xls.Add(fi);  
				}
                    
				////
				///  Recurse the subdirectories
				///
				Visit(diNext.FullName+"\\",Filter,first_time);
 
			}

		}

	}

}
