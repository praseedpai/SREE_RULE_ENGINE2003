using System;
using System.Collections;
using System.IO;
using System.Xml; 

namespace ForeCastWatcher
{
	/// <summary>
	///  Token Enumerations for Communication between Lexer and 
	///  Parsers
	/// </summary>
	public enum TOKEN
	{
		ILLEGAL_TOKEN=-1, // Not a Token
		TOK_COMMENT=1,      // Comment Token ( presently not used )   
		TOK_NULL,         // END OF INPUT
		TOK_STOP,         // STOP STATEMENT
		TOK_LITERAL,      // LITERAL STRING "<string>"
		TOK_VALIDATE,     // VALIDATE STATEMENT
		TOK_BUILTIN_SUM_RANGE, // SUM A RANGE OF CELLS
		TOK_BUILTIN_SUM_COLUMN,// SUM A SPREAD SHEET TABLE COLUMN 
		TOK_BUILTIN_LOOKUP,    // Lookup a cell value in DB
		TOK_BUILTIN_VAL,       // Convert a string to value 
		TOK_BOOL_TRUE,         // Boolean TRUE
		TOK_BOOL_FALSE,        // Boolean FALSE
		TOK_PLUS,              // '+'
		TOK_MUL,               // '*' 
		TOK_DIV,               // '/'
		TOK_SUB,               // '-'
		TOK_OPAREN,            // '(' 
		TOK_CPAREN,            // ')' 
		TOK_DOUBLE,            // [0-9]+ 
		TOK_EQ,                // '=='
		TOK_NEQ,               // '<>'
		TOK_GT,                // '>'
		TOK_GTE,               // '>='
		TOK_LT,                // '<'
		TOK_LTE,               // '<='
		TOK_AND,               // '&&'
		TOK_OR,                // '||'
		TOK_NOT,               // '!'
		TOK_ASSIGN,            // '='
		TOK_CELL,              // Spreasheet cell 
		TOK_SEMI,              // ';' 
		TOK_DBCOLUMN,          // DB column  
		TOK_COMMA,             // ','  
		TOK_STRING,            // [A-Za-z][A-za-z_0-9]*
		TOK_HASH,              // '#'
		TOK_CELLDBCOLUMN,      // Spreadsheet cell as DB Column 
		TOK_IF,                // IF 
		TOK_WHILE,             // WHILE
		TOK_VAR_NUMBER,        // NUMBER data type
		TOK_VAR_STRING,        // STRING data type 
		TOK_VAR_BOOL,          // Bool data type
		TOK_SET,               // Set Statement ( not implemented) 
		TOK_WEND,              // Wend Statement
		TOK_ELSE,              // Else Statement
		TOK_ENDIF,             // Endif Statement
		TOK_UNQUOTED_STRING,   // Variable name 
		TOK_COLON,             // ':' ( not used )
		TOK_VAR,               // Token for Variable
		TOK_THEN,              // THEN Keyword
		TOK_END,               // END keyword
		TOK_WRITELOG,          // Log 
		TOK_BEGIN,             // BEGIN KeyWord 
		TOK_EXCEPTION,         // Exception
		TOK_SCAN,              // SCAN statement
		TOK_SCAN_PULL,         // SCAN pull statement 
		TOK_CHECKDUP,          // CheckDup statement
		TOK_CELLNAME,          // CellName
		TOK_CLEANUP,           // Cleanup
		TOK_CHECKHEADER,       // Check whether a genuine heading exits for worksheet
		TOK_SCAN_BRANDID,      // Scan the Brand id   
		TOK_CHECKCURRENCY,     // Check for VALID Currency
		TOK_RECORD_COUNT,      // Record Count  
		TOK_UPPER,             // UPPER
		TOK_TRIM,                   // TRIM   
		TOK_CHECK_HOUR,
		TOK_BUILTIN_LOOKUP_CASE_IGNORE ,
		TOK_CHECKNULL
}

	/// <summary>
	///    class for Exception Handling 
	/// </summary>
	public class CParserException : Exception
	{
		private int ErrorCode ;
		private String ErrorString;
		private int Lexical_Offset;
		/// <summary>
		///   Ctor
		/// </summary>
		/// <param name="pErrorCode"></param>
		/// <param name="pErrorString"></param>
		/// <param name="pLexical_Offset"></param>

		public CParserException( int pErrorCode ,
			String pErrorString,
			int pLexical_Offset)
		{
			ErrorCode   = pErrorCode;
			ErrorString = pErrorString;
			Lexical_Offset = pLexical_Offset;
		}
		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public int GetErrorCode()
		{
			return ErrorCode;
		}
		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public String GetErrorString()
		{
			return ErrorString; 
		}
		/// <summary>
		///   
		/// </summary>
		/// <returns></returns>

		public int GetLexicalOffset()
		{
			return Lexical_Offset; 
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="lex"></param>

		public void SetLexicalOffset(int lex)
		{
			Lexical_Offset = lex;
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="pStr"></param>

		public void SetErrorString(String pStr)
		{
			ErrorString = pStr;
		}


	}
	/// <summary>
	///    Error for semntic erros
	/// </summary>
	/// 
	class CSemanticErrorLog 
	{
		/// <summary>
		///     
		/// </summary>
		static int ErrorCount=0;
		static ArrayList lst=new  ArrayList();
		/// <summary>
		///    Ctor
		/// </summary>
 		 static CSemanticErrorLog()
		{
			
		}

        /// <summary>
        /// 
        /// </summary>
		public static void Cleanup()
		{
			lst.Clear(); 
			ErrorCount = 0;   
		}
		/// <summary>
		///    Get Logged data as a String 
		/// </summary>
		/// <returns></returns>
		public static String GetLog()
		{


			String str="Logged data by the user and processing status"+"\r\n";
			str += "--------------------------------------\r\n";  

			int xt = lst.Count;

			if ( xt == 0 )
			{
				str += "NIL"+"\r\n";
               
			}
			else 
			{

				for(int i=0; i < xt; ++i )
				{
					str = str + lst[i].ToString()+"\r\n"; 
				}
			}
			str += "--------------------------------------\r\n";  
			return str;
		}
		/// <summary>
		///    Add a Line to Log
		/// </summary>
		/// <param name="str"></param>
		public static void AddLine( String str )
		{
			lst.Add(str.Substring(0));   
			ErrorCount++;
		}
		/// <summary>
		///   Add From a Script   
		/// </summary>
		/// <param name="str"></param>

		public static void AddFromUser(String str )
		{
			lst.Add(str.Substring(0));   
			ErrorCount++;
         
		}


      
	}
	/// <summary>
	///    
	/// </summary>
    class CSyntaxErrorLog
	{

		/// <summary>
		///   instance variables
		/// </summary>
		static int ErrorCount=0;
		static ArrayList lst = new ArrayList() ;
        /// <summary>
        ///    Ctor
        /// </summary>
		static CSyntaxErrorLog()
		{
			
		}


		public static void Cleanup()
		{
               lst.Clear(); 
               ErrorCount = 0;   
		}
		/// <summary>
		///    Add a Line from script
		/// </summary>
		/// <param name="str"></param>

		public static void AddLine( String str )
		{
			lst.Add(str.Substring(0));   
			ErrorCount++;

		}

		/// <summary>
		///    Get Logged data as a String 
		/// </summary>
		/// <returns></returns>
		public static String GetLog()
		{
			
			String str="Syntax Error"+"\r\n";
			str += "--------------------------------------\r\n";  

			int xt = lst.Count;

			if ( xt == 0 )
			{
				str += "NIL"+"\r\n";
               
			}
			else 
			{

				for(int i=0; i < xt; ++i )
				{
					str = str + lst[i].ToString()+"\r\n"; 
				}
			}
			str += "--------------------------------------\r\n";  
			return str;
		}
	}


	class GlobalSettings
	{

		/// <summary>
		///     State variables
		/// </summary>
		public static int  num_decimal_places=1;
		public static bool allow_emails=false;
		public static bool backup_sheets=false;
		public static bool quarantine_sheets=false;
		public static bool event_logging=false;
		public static bool request_processing_thread_abort=false;
		public static string xls_file_path;
		public static int oracle_connection_count = 0;
		public static int excel_connection_count = 0;
		

		
		/// <summary>
		///     To Store The Values in a Hash map
		/// </summary>
		static Hashtable config_table = new System.Collections.Hashtable();
		/// <summary>
		///    Static constructor meant for intialization 
		/// </summary>
		static GlobalSettings()
		{
            xls_file_path ="UNKNOWN";
			
            
    
		}

		////////////////////////////////////////////
		///
		///
		///
		public static void Cleanup()
		{

			if ( config_table != null )
             config_table.Clear();
 


		}


		
		////////////////////////////////////////////
		///
		///
		///
		public static void AddEntry(String str, Object obj )
		{
			try 
			{
				config_table.Add(str,obj); 

				if ( str == "DECIMAL_ROUNDING" )
				{
					num_decimal_places = Convert.ToInt16(obj); 
					if ( num_decimal_places > 5 )
						num_decimal_places = 5;
					else if ( num_decimal_places < 0 )
						num_decimal_places = 1;


				}
				else if ( str == "BACKUP_SHEETS" )
				{
					String test = Convert.ToString(obj).Trim().ToUpper();
  
					if ( test == "YES")
					{
						backup_sheets = true;

					}
					else 
					{
						backup_sheets = false; 

					}

				}
				else if ( str == "QUARANTINE_SHEETS")
				{

					String test = Convert.ToString(obj).Trim().ToUpper();
  
					if ( test == "YES")
					{
						quarantine_sheets = true;

					}
					else 
					{
						quarantine_sheets = false; 

					}

                   
				}
				else if ( str == "ALLOW_EMAILS")
				{

					String test = Convert.ToString(obj).Trim().ToUpper();
  
					if ( test == "YES")
					{
						allow_emails = true;

					}
					else 
					{
						allow_emails = false; 

					}

                   
				}
				else if ( str == "EVENT_LOGGING")
				{

					String test = Convert.ToString(obj).Trim().ToUpper();
  
					if ( test == "YES")
					{
						event_logging = true;

					}
					else 
					{
						event_logging = false; 

					}

				}
			}
			catch(Exception  )
			{

                // Error while updating the data


			}

                   
			



		}
		////////////////////////////////////////////////
		///  Get The entry
		///
		///
		public static object GetEntry(String str )
		{
            return config_table[str]; 

		}
		////////////////////////////////////////////////////
		///
		///
		///
		///
        
	}
	/// <summary>
	///  Type info enumerations
	/// </summary>
	public enum TYPE_INFO
	{
		TYPE_ILLEGAL = -1, // NOT A TYPE
		TYPE_DOUBLE ,      // IEEE Double precision floating point 
		TYPE_BOOL,         // Boolean Data type
		TYPE_STRING,       // String data type 
		TYPE_DATE          // Date Data Type 
	}
	/// <summary>
	///    Symbol Table entry for variable
	/// </summary>
	public class Symbol
	{
		public String SymbolName;   // Symbol Name
		public TYPE_INFO Type;      // Data type
		public String str_val;      // memory to hold string 
		public double dbl_val;      // memory to hold double
		public bool   bol_val;      // memory to hold boolean
		public DateTime date_val;   // Date Time field 
	}
	/// <summary>
	///  Label info class
	///  (not used )
	/// </summary>
	public class CLabel
	{
		public String LabelName;   // Label name 
		public int    SrcOffset;   // Offset in the source

	}
	/// <summary>
	///     Keyword Table Entry
	/// </summary>
	/// 
	public struct ValueTable
	{
		public TOKEN tok;          // Token id
		public String Value;       // Token string  
		public ValueTable( TOKEN tok , String Value )
		{
			this.tok = tok;
			this.Value = Value;
         
		}
	}
	/// <summary>
	///  Data structure to model Spreadsheet Cell reference
	/// </summary>
	public struct CellValue
	{
		public String SheetName;  // Spread sheet name
		public String CellName;   // Cellname

		/// <summary>
		///    Ctor
		/// </summary>
		/// <param name="s"></param>
		/// <param name="c"></param>
		public CellValue(String s, String c )
		{
			SheetName = s.ToUpper();
			CellName  = c.ToUpper();
		}
	}
	/// <summary>
	///   DB column reference
	/// </summary>
	public struct CTableColumn
	{
		public String TableName;  // Table name
		public String ColumnName; // Coumn name
  
		/// <summary>
		///    Ctor
		/// </summary>
		/// <param name="s"></param>
		/// <param name="c"></param>
		public CTableColumn(String s, String c )
		{
			TableName   = s.Substring(0);
			ColumnName  = c.Substring(0);
		}   
	}
	/// <summary>
	///   Treat spread sheet cell as a column
	/// </summary>
	public struct CellAsTableColumn
	{
		public String SheetName;  // Sheetname
		public String CellName;   // Cellname
		public int    Index;      // index number of the record
		/// <summary>
		///    Ctor
		/// </summary>
		/// <param name="s"></param>
		/// <param name="c"></param>
		/// <param name="i"></param>

		public CellAsTableColumn(String s, String c , int i)
		{
			SheetName = s;
			CellName = c;
			Index=i;

		}

	}
		
	/// <summary>
	///  Symbol Table Class. Collection of items.
	/// </summary>
	public class SymbolTable
	{

		private Hashtable ar;

		/// <summary>
		///   Ctor
		/// </summary>
		public SymbolTable()
		{
			ar = new Hashtable();
		}
		/// <summary>
		///    Add an entry into symbol table
		/// </summary>
		/// <param name="key"></param>
		/// <param name="Value"></param>

		public void Add( object key , object Value )
		{
			ar.Add(key,Value);
	
		}
		/// <summary>
		///     Find the value corresponding to a key
		/// </summary>
		/// <param name="key"></param>
		/// <returns></returns>
		public object Lookup(object key)
		{
			return ar[key];
		}
		/// <summary>
		///  
		/// </summary>

		public void Clear()
		{
              ar.Clear();
               

		}
	
	}

	/// <summary>
	///   A push down stack for double precision floating points
	/// </summary>
	/// 
	public class Stack
	{
	
		long[] stk;             // contents are stored here   
		int top_stack ;         // stack pointer
		/// <summary>
		///   Ctor
		/// </summary>
		public Stack()
		{
			stk = new long[1024];
			top_stack = 0;
		}
		/// <summary>
		///  Empty the stack before evaluation of expression
		/// </summary>
		public void Clear()
		{
			top_stack = 0;

		}
		/// <summary>
		///    Extract the bit patterns of a double value ( high dword )
		/// </summary>
		/// <param name="d"></param>
		/// <returns></returns>
		unsafe private long HIGH_DWORD_FROM_DOUBLE(double d)
		{
			long *x = (long *)(void *)&d;
			return x[1];    

		}
		/// <summary>
		///   Extract the bit patterns of a double value (low dword )
		/// </summary>
		/// <param name="d"></param>
		/// <returns></returns>
		unsafe private long LOW_DWORD_FROM_DOUBLE(double d)
		{
			long *x = (long *)(void *)&d;
			return x[0];     
		}

		/// <summary>
		///    Push the value into the stack
		/// </summary>
		/// <param name="dbl"></param>
   
		public void push(long  dbl)
		{
			if ( top_stack == 1023 )
			{
				CSyntaxErrorLog.AddLine("Stack OverFlow"); 
				throw new CParserException(-100,"Stack OverFlow",-1);
			}
			stk[top_stack++] = dbl;
		}
		/// <summary>
		///    Push a IEEE 754 double into stack
		/// </summary>
		/// <param name="dbl"></param>

		public void push(double dbl )
		{

			if ( top_stack == 1023 )
			{
				CSyntaxErrorLog.AddLine("Stack OverFlow"); 
				throw new CParserException(-100,"Stack OverFlow",-1);
			}
			stk[top_stack++] = LOW_DWORD_FROM_DOUBLE(dbl);
			stk[top_stack++] = HIGH_DWORD_FROM_DOUBLE(dbl); 
		}
		/// <summary>
		///   Pop a value from the stack
		/// </summary>
		/// <returns></returns>
		public long  pop()
		{
			if ( top_stack == 0 )
			{
				CSyntaxErrorLog.AddLine("Stack Underflow"); 
				throw new CParserException(-100,"Stack Underflow",-1);
			}
			return stk[--top_stack];
		}
		/// <summary>
		///    Method which retrives two long values and combine to form
		///    a IEEE 754 floating point value 
		/// </summary>
		/// <returns></returns>

		unsafe private double PopWorker()
		{

			long hd = pop();
			long ld = pop();
			fixed(long *dbfr = new long[2])
			{
				dbfr[0]=ld;
				dbfr[1]=hd;
				return * ( (double *) (void *) &dbfr[0]); 
			}
		}
		/// <summary>
		///  Pop a IEEE 754 double precision floating point value from the stack
		/// </summary>
		/// <returns></returns>
		public double PopD()
		{
			return PopWorker();
              
		}
		/// <summary>
		///     To Check Whether the stack is empty
		/// </summary>
		/// <returns></returns>
		public bool IsEmpty()
		{
			return ( top_stack == 1 ) ? true : false;
		}

		/// <summary>
		///    Check whether we can pop a double value from the stack
		/// </summary>
		/// <returns></returns>
 
		public bool CanPopDouble()
		{
			return ( top_stack == 2 ) ? true : false;    

		}
         
	}

	
	/// <summary>
	///   Parser For Script 
	/// </summary>
	public class CParser : Lexer
	{
		Stack  ValueStack = new Stack();    // Stack to Aid expression evaluation
		ArrayList StringList = new ArrayList(); //String pool
		
		
		/// <summary>                                  
		///  Symbol Table for the APP
		/// </summary>
		SymbolTable sym = new SymbolTable(); 

		String EmailToPeople="";

		String CountryFromFolder = "";
		String SegmentFromFolder = "";

		
		/// <summary>
		///   Validate Boolean Expressions
		/// </summary>
		/// 
		public TYPE_INFO ValidateStatement()
		{
			
			GetNext();

			// Evaluate the expression
			object c =CallExpr();
             
			if ( c.GetType() != Type.GetType("System.Boolean"))
			{

                 CSyntaxErrorLog.AddLine("Boolean Value expected at the following line");     
				 CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				 throw new CParserException(-100,"Boolean Expression expected",-1); 
                   
			}


			if ( Current_Token != TOKEN.TOK_SEMI )
			{
				CSyntaxErrorLog.AddLine("Semi Colon expected at the following line"+"\r\n");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Semi Colon Expected",-1); 
               

			}

			bool ct = (bool)c;


			///////////////////////////////////////
			///
			/// Exception code with error code of 100 is to
			/// implement BEGIN....EXCEPTION...END Block
			///
			if ( ct == false )
			{
				throw new CParserException(100,
					"Validation Failed",
					this.SaveIndex());  

			}

			return TYPE_INFO.TYPE_BOOL; 
		}
		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>

		public String GetSyntaxLog()
		{

			return CSyntaxErrorLog.GetLog();  

		}
		/// <summary>
		///    Get the Email of the Person who created the excel file
		/// </summary>
		/// <returns></returns>
		public String GetEmail()
		{

			return EmailToPeople;
 
			
                


		}

		/// <summary>
		///    Get Excel reader instance
		/// </summary>
		/// <returns></returns>
		public CExcelReader GetExcelReader()
		{
			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");
			return st;
           
		}
		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public String GetSemanticLog()
		{
    		return CSemanticErrorLog.GetLog();  
		}
    	/// <summary>
		///    
		/// </summary>

		public void StopStatement()
		{
			GetNext();
			ValueStack.Clear();
            TYPE_INFO RetValue = BExpr(); 

			if ( RetValue != TYPE_INFO.TYPE_DOUBLE )
			{

				CSyntaxErrorLog.AddLine("A Numeric value expected as return code");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"A numeric error code expected",-1);

			}

			double nd = ValueStack.PopD();  

			if ( Current_Token != TOKEN.TOK_SEMI )
			{
				CSyntaxErrorLog.AddLine("Semi Colon expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,
					"Semi Colon Expected",
					this.SaveIndex()); 

			}

			if ( nd == 0.0 ) 
			{
			    
				throw new CParserException(200,"Successful Validation",-1);
			}
			else 
			{
				string err_code = Convert.ToString(nd); 
				CSyntaxErrorLog.AddLine("Validation returned with errorcode= "+err_code);     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200,"Validation Failure",-1);
			}
 
		}
		/// <summary>
		///   Sum Builtin function
		/// </summary>
		public TYPE_INFO ParseSum()
		{
			
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{
				CSyntaxErrorLog.AddLine("( expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Parenthesis expected",-1);
                 
			}

			GetNext();

			if ( Current_Token != TOKEN.TOK_CELL)
			{
				
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
			}

			CellValue first_cell = new CellValue(base.cell_value.SheetName ,
				base.cell_value.CellName );  
			// Extracted Value 
			GetNext();

			if ( Current_Token != TOKEN.TOK_COMMA)
			{
				CSyntaxErrorLog.AddLine(", expected in the following line");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," , expected",-1);
			}
              
              
			GetNext(); 

			if ( Current_Token != TOKEN.TOK_CELL)
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
			}

			CellValue second_cell = new CellValue(base.cell_value.SheetName ,
				base.cell_value.CellName );  
			// Extracted Value 

			GetNext();
  

			// Extracted Value
			if ( Current_Token != TOKEN.TOK_CPAREN )
			{
				CSyntaxErrorLog.AddLine(") Paren expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," ) expected",-1);
                 
			}
         
			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");

			
			double ret = st.SumRange(first_cell.SheetName ,first_cell.CellName ,
				second_cell.CellName ,100);

			ValueStack.push(ret); 
			
			GetNext();

			return TYPE_INFO.TYPE_DOUBLE;  

		}
		/// <summary>
		///    Write the Log
		/// </summary>

		public void WriteLogStatement()
		{
			GetNext();

			ValueStack.Clear();

			TYPE_INFO RetValue = BExpr(); 

			if ( RetValue != TYPE_INFO.TYPE_STRING )
			{
				CSemanticErrorLog.AddFromUser("LOG ERROR"); 
				
			}
			else 
			{
				int nd = (int)ValueStack.pop();
				CSemanticErrorLog.AddFromUser(StringList[nd].ToString());
               
			}

			if ( Current_Token != TOKEN.TOK_SEMI )
			{

				throw new CParserException(-100,
					"Semi Colon Expected in Validate Statement",
					this.SaveIndex()); 

			}

		

			return; 
           

		}
		/// <summary>
		/// 
		/// </summary>
		/// <returns></returns>
		public TYPE_INFO ScanStatement()
		{

			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{
				CSyntaxErrorLog.AddLine("( Paren expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100," ( expected",-1);
		      
			}


			GetNext();

			if ( Current_Token != TOKEN.TOK_CELLNAME )
			{
				CSyntaxErrorLog.AddLine("Invalid cell name");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100," Invalid cell name",-1);
			}

			CellValue x = new CellValue(base.cell_value.SheetName ,
				base.cell_value.CellName );
			// Extracted Value 
			GetNext();

			if ( Current_Token != TOKEN.TOK_COMMA)
			{
				CSyntaxErrorLog.AddLine(", expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100," , expected",-1);
			}

              
			GetNext(); 
			if ( Current_Token != TOKEN.TOK_DBCOLUMN)
			{
				
				CSyntaxErrorLog.AddLine(" Invalid DB Column");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100," Invalid DB Column",-1);
				
			}

			CTableColumn ct = new CTableColumn(base.tbl_clm.TableName ,
				base.tbl_clm.ColumnName ); 
			// Extracted Value 
			GetNext();

			// Extracted Value
			if ( Current_Token != TOKEN.TOK_CPAREN )
			{
				CSyntaxErrorLog.AddLine("Closing Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Closing Parenthesis expected",-1);
				
                 
			}

			/////////////////////////////////////////
			///
			///
			///
			///
			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");
 
			bool rs = st.ScanBySendingExcelData(x.SheetName ,x.CellName ,ct.TableName , ct.ColumnName ); 
			
			ValueStack.push((long)(( rs == true )? 1 : 0 ));
			GetNext();
			return TYPE_INFO.TYPE_BOOL; 

			

		}

		/// <summary>
		///    Scan pull statement
		/// </summary>
		/// <returns></returns>


		public TYPE_INFO ScanPullStatement()
		{

			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{
				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
					
                
			}

			GetNext();

			if ( Current_Token != TOKEN.TOK_CELLNAME )
			{
				CSyntaxErrorLog.AddLine("Invalid Cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Invalid Cell reference",-1);
		
			}

			CellValue x = new CellValue(base.cell_value.SheetName ,
				base.cell_value.CellName );
			// Extracted Value 
			GetNext();

			if ( Current_Token != TOKEN.TOK_COMMA)
			{
				CSyntaxErrorLog.AddLine(", expected in the following line");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," , expected",-1);
				
			}

              
			GetNext(); 
			if ( Current_Token != TOKEN.TOK_DBCOLUMN)
			{
				CSyntaxErrorLog.AddLine(" Invalid DB Column");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100," Invalid DB Column",-1);
				
			}

			CTableColumn ct = new CTableColumn(base.tbl_clm.TableName ,
				base.tbl_clm.ColumnName ); 
			// Extracted Value 
			GetNext();

			// Extracted Value
			if ( Current_Token != TOKEN.TOK_CPAREN )
			{
				CSyntaxErrorLog.AddLine("Closing Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Closing Parenthesis expected",-1);
				
                 
			}

			/////////////////////////////////////////
			///
			///
			///
			///
			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");
 
			bool rs = st.ScanInDB(x.SheetName ,x.CellName ,ct.TableName , ct.ColumnName ); 

			ValueStack.push((long)(( rs == true )? 1 : 0 ));
			GetNext();
			return TYPE_INFO.TYPE_BOOL; 



 
		}

#if false
		/// <summary>
		///   Sum Builtin function
		/// </summary>
		public TYPE_INFO ParseSumColumn()
		{
			GetNext();
			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{
				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
                 
			}

			GetNext();

			if ( Current_Token != TOKEN.TOK_CELLDBCOLUMN)
			{
				CSyntaxErrorLog.AddLine("Invalid spreadsheet reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Invalid Spreadsheet cell reference",-1);
				
			}

			String str_name   = base.cell_tbl_clm.SheetName ;
			String str_column = base.cell_tbl_clm.CellName ;
			GetNext();
			TYPE_INFO RetValue = BExpr();

			if ( RetValue != TYPE_INFO.TYPE_DOUBLE )
			{
				CSyntaxErrorLog.AddLine("Numeric Value expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Numeric Value Expected",-1);
				
			}
			double temp  = ValueStack.PopD();
			base.cell_tbl_clm.Index = ( int)temp; 

			CellAsTableColumn first_cell = new CellAsTableColumn(
				str_name ,
				str_column,
				(int)temp);  
			// Extracted Value 
			GetNext();

			if ( Current_Token != TOKEN.TOK_COMMA)
			{
				CSyntaxErrorLog.AddLine(", expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100," , expected",-1);
			}
			GetNext();

			if ( Current_Token != TOKEN.TOK_CELLDBCOLUMN)
			{
				throw new Exception("Invalid Spreadsheet cell reference");
			}
			str_name   = base.cell_tbl_clm.SheetName ;
			str_column = base.cell_tbl_clm.CellName ;
			GetNext();
			RetValue = BExpr();

			if ( RetValue != TYPE_INFO.TYPE_DOUBLE )
				throw new Exception("Long type Expected");
			temp  = ValueStack.PopD();
			base.cell_tbl_clm.Index = ( int)temp; 

			GetNext();
            
			if ( Current_Token != TOKEN.TOK_CPAREN)
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
				
			}

			CellAsTableColumn second_cell = new CellAsTableColumn(
				str_name ,
				str_column,
				(int)temp);  
			// Extracted Value 

			GetNext();
  

			         
			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");

			
			double ret = st.SumColumn(first_cell.SheetName ,first_cell.CellName ,
				first_cell.Index ,second_cell.Index );

			
			ValueStack.push(Math.Round(ret*100)); 
			
			
			return TYPE_INFO.TYPE_DOUBLE;  
		}

#endif



		/// <summary>
		///   
		/// </summary>
		/// <returns></returns>
		public TYPE_INFO ParseScanBrand()
		{
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{

				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
				
                 
			}
			GetNext();

			if ( Current_Token != TOKEN.TOK_CPAREN  )
			{

				CSyntaxErrorLog.AddLine("Closing Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,") Parenthesis expected",-1);
              
			}

			////////////////////////////////////////////////////
			///
			///
			///

            CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");

 			bool rs = st.ScanBrand(); 

			ValueStack.push((long)(( rs == true )? 1 : 0 ));

			GetNext();

    		return TYPE_INFO.TYPE_BOOL;
 
		}



		/// <summary>
		///     
		/// </summary>
		/// <returns></returns>

		public TYPE_INFO ParseCheckHourly()
		{
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{

				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
				
                 
			}

			GetNext();


			if ( Current_Token != TOKEN.TOK_CPAREN  )
			{

				CSyntaxErrorLog.AddLine("Closing Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,") Parenthesis expected",-1);
				
                 
			}

    		CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");
			bool rs = st.CheckHourly(); 
			ValueStack.push((long)(( rs == true )? 1 : 0 ));
		    GetNext();
			return TYPE_INFO.TYPE_BOOL;


		
 
		}


		/// <summary>
		///     Val Function
		/// </summary>
		/// <returns></returns>

		public TYPE_INFO ParseIsCurrency()
		{
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{

				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
				
                 
			}

			TYPE_INFO RetValue = BExpr();

			if ( RetValue != TYPE_INFO.TYPE_STRING  )
			{
				CSyntaxErrorLog.AddLine("String Value expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"String Value Expected",-1);
				
			}
 
			int x = (int)ValueStack.pop(); 



			String Str =(String) StringList[x];


			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");

			bool rs = st.ScanCurrency(Str); 

			ValueStack.push((long)(( rs == true )? 1 : 0 ));

			
			return TYPE_INFO.TYPE_BOOL;




			
 
		}


		/// <summary>
		///     Val Function
		/// </summary>
		/// <returns></returns>

		public TYPE_INFO ParseRecCount()
		{
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{

				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
				
                 
			}

			TYPE_INFO RetValue = BExpr();

			if ( RetValue != TYPE_INFO.TYPE_STRING  )
			{
				CSyntaxErrorLog.AddLine("String Value expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"String Value Expected",-1);
				
			}
 
			int x = (int)ValueStack.pop(); 



			String Str =(String) StringList[x];


			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");

			double  rec_count = st.RecCount(Str); 

			ValueStack.push(rec_count);

			return TYPE_INFO.TYPE_DOUBLE;




			
 
		}

		/// <summary>
		///     TRIM Function
		/// </summary>
		/// <returns></returns>

		public TYPE_INFO ParseTrim()
		{
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{

				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
				
                 
			}

			TYPE_INFO RetValue = BExpr();

			if ( RetValue != TYPE_INFO.TYPE_STRING  )
			{
				CSyntaxErrorLog.AddLine("String Value expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"String Value Expected",-1);
			}
 
			int x = (int)ValueStack.pop(); 
			String Str =(String) StringList[x];
			Str = Str.Trim();
			int newstr = StringList.Add(Str); 
			ValueStack.push(newstr);  
			return TYPE_INFO.TYPE_STRING;

		}

		/// <summary>
		///    Upper
		/// </summary>
		/// <returns></returns>


		public TYPE_INFO ParseUpper()
		{
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{

				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
				
                 
			}

			TYPE_INFO RetValue = BExpr();

			if ( RetValue != TYPE_INFO.TYPE_STRING  )
			{
				CSyntaxErrorLog.AddLine("String Value expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"String Value Expected",-1);
			}
 
			int x = (int)ValueStack.pop(); 
			String Str =(String) StringList[x];
			Str = Str.Trim();
			int newstr = StringList.Add(Str); 
			ValueStack.push(newstr);  
			return TYPE_INFO.TYPE_STRING;

		}
		/// <summary>
		///     Val Function
		/// </summary>
		/// <returns></returns>

		public TYPE_INFO ParseVal()
		{
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{

				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
				
                 
			}

			TYPE_INFO RetValue = BExpr();

			if ( RetValue != TYPE_INFO.TYPE_STRING  )
			{
				CSyntaxErrorLog.AddLine("String Value expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"String Value Expected",-1);
				
			}
 
			int x = (int)ValueStack.pop(); 



			String Str =(String) StringList[x];

			if ( Str == "" )
			{
				ValueStack.push(0);  
			}
			else 
			{
				double rs = ConvertStringToValue(Str);
				ValueStack.push(rs);   
			}


			
			
			return TYPE_INFO.TYPE_DOUBLE;
 
		}
		/// <summary>
		///   Lookup Builtin Function
		/// </summary>

		public TYPE_INFO ParseLookup(bool ignorecase)
		{
			GetNext();
			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{
				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
                 
			}
			GetNext();

			if ( Current_Token != TOKEN.TOK_CELL )
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
			}

			CellValue x = new CellValue(base.cell_value.SheetName ,
				base.cell_value.CellName );
			// Extracted Value 
			GetNext();

			if ( Current_Token != TOKEN.TOK_COMMA)
			{
				CSyntaxErrorLog.AddLine(", expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100," , expected",-1);
			}

              
			GetNext(); 
			if ( Current_Token != TOKEN.TOK_DBCOLUMN)
			{
				CSyntaxErrorLog.AddLine(" Invalid DB Column");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100," Invalid DB Column",-1);
				
			}

			CTableColumn ct = new CTableColumn(base.tbl_clm.TableName ,
				base.tbl_clm.ColumnName ); 
			// Extracted Value 
			GetNext();

			// Extracted Value
			if ( Current_Token != TOKEN.TOK_CPAREN )
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
                 
			}

			/////////////////////////////////////////
			///
			///
			///
			///
			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");
 

			bool rs;
			
			if ( ignorecase )
			{
				rs = st.LookupInDB(x.SheetName ,x.CellName ,ct.TableName , ct.ColumnName ); 
			}
			else 
			{
                rs = st.LookupInDBCaseSensitive(x.SheetName ,x.CellName ,ct.TableName , ct.ColumnName ); 

			}
			ValueStack.push((long)(( rs == true )? 1 : 0 ));
			GetNext();
			return TYPE_INFO.TYPE_BOOL; 

		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="str"></param>
		/// <returns></returns>
		public double ConvertStringToValue(String str)
		{
			String ret_value = "";
			bool is_dec = false;
			bool pos= true;
			int index = 0;  

			str = str.Trim();
 
			if (!(char.IsDigit(str[index]) || str[index] != '.' ))
				               return 0.0;  

			if (str[index] == '-' || str[index] == '+' ) 
			{
				pos = (str[index] == '+') ? true:false;
				index++;
			}

			

			while (index < str.Length && (char.IsDigit(str[index]) || str[index] == '.' ))
			{
				if ( str[index] == '.'  && is_dec == false )
					is_dec = true;
				else if (( str[index] == '.'  && is_dec == true  ))
					return 0.0;
                   
				ret_value = ret_value + str[index];
				index++;
			}
            
			if ( ret_value.Trim() == "" )
				return 0.0;
			else 
			{
				double ret_val = Convert.ToDouble(ret_value); 
				if ( !pos )
					ret_val = -ret_val;
				return ret_val;
			}

		}
        
		/// <summary>
		///   Open XLS file ( Workbook name is implicit
		/// </summary>
		public TYPE_INFO OpenStatement(String ExcelConStr,String OrclConStr)
		{
			try 
			{
				sym.Add("EXCEL_READER",new CExcelReader(ExcelConStr,OrclConStr));
				CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");
				EmailToPeople = st.GetCellValue("CONTROL","B10");

				
				String country = st.GetCellValue("CONTROL","B2");
				String segment  = st.GetCellValue("CONTROL","B3");

				country = country.Trim().ToUpper();
				segment = segment.Trim().ToUpper(); 

				if ( CountryFromFolder != country  || 
					SegmentFromFolder != segment )
				{
                    CSemanticErrorLog.AddLine("Country/Segment Folder does not match the Country and Segment in Worksheet");    
					throw new CParserException(-1,"Country/Segment Folder does not match the Country and Segment in Worksheet",-1);
					 

				}
					  



			}
			catch(CParserException et )
			{
                 
				 throw et;

			}
			catch(Exception e )
			{
				CSyntaxErrorLog.AddLine("Excel File cannot be opened"+e.ToString() );      
                  
			}

			return TYPE_INFO.TYPE_BOOL; 
		}
		/// <summary>
		///     Begin ...Exception....End Block
		/// </summary>
		void BeginStatement()
		{
			GetNext();
			try 
			{
				StatementList();
				if ( Current_Token == TOKEN.TOK_EXCEPTION )
				{
					GetNext();
					SkipToEnd(); 
				}

			}
			catch(CParserException e )
			{
				if ( e.GetErrorCode() == 100 )
				{
					SkipToException();
					GetNext();
					StatementList();
				}

			}

		}


		/// <summary>
		///     Check for Null
		/// </summary>
		public TYPE_INFO ParseCheckNull()
		{
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{
				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);

                 
			}

			GetNext();

			if ( Current_Token != TOKEN.TOK_CELLNAME)
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
			}

			ArrayList arr = new ArrayList(); 
			String SheetName;

			SheetName = base.cell_value.SheetName;
 
			while ( Current_Token == TOKEN.TOK_CELLNAME )  
			{
				arr.Add(base.cell_value.CellName); 
				// Extracted Value 
				GetNext();
				if ( Current_Token == TOKEN.TOK_COMMA)
				{
					GetNext(); 
				}
				
			}

			if ( Current_Token != TOKEN.TOK_CPAREN )
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
			}
			         
			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");
			bool  ret = st.HasNull(SheetName,arr);
			ValueStack.push((long)((ret)?1:0)); 			
			GetNext();
			return TYPE_INFO.TYPE_BOOL;  
		}

		/// <summary>
		///     Check for Duplication
		/// </summary>
		public TYPE_INFO ParseCheckDup()
		{
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{
				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);

                 
			}

			GetNext();

			if ( Current_Token != TOKEN.TOK_CELLNAME)
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
			}

			ArrayList arr = new ArrayList(); 
			String SheetName;

			SheetName = base.cell_value.SheetName;
 
			while ( Current_Token == TOKEN.TOK_CELLNAME )  
			{
				arr.Add(base.cell_value.CellName); 
				// Extracted Value 
				GetNext();
				if ( Current_Token == TOKEN.TOK_COMMA)
				{
					GetNext(); 
				}
				
			}

			if ( Current_Token != TOKEN.TOK_CPAREN )
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
			}

			         
			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");

			bool  ret = st.CheckDuplicate(SheetName,arr);

			ValueStack.push((long)((ret)?1:0)); 			
			GetNext();
			return TYPE_INFO.TYPE_BOOL;  
		}
		/// <summary>
		///    Clean up cells
		/// </summary>
		public TYPE_INFO ParseCleanup()
		{
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{
				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
                 
			}

			GetNext();

			if ( Current_Token != TOKEN.TOK_CELLNAME)
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
			}

			ArrayList arr = new ArrayList(); 
			String SheetName;

			SheetName = base.cell_value.SheetName;
 
			while ( Current_Token == TOKEN.TOK_CELLNAME )  
			{
				arr.Add(base.cell_value.CellName); 
				// Extracted Value 
				GetNext();
				if ( Current_Token == TOKEN.TOK_COMMA)
				{
					GetNext(); 
				}
				
			}

			if ( Current_Token != TOKEN.TOK_CPAREN )
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
			}

			         
			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");

			bool  ret = st.CleanUpCells(SheetName,arr);

			ValueStack.push((long)((ret)?1:0)); 			
			GetNext();
			return TYPE_INFO.TYPE_BOOL;  
		}


		/// <summary>
		///    Clean up cells
		/// </summary>
		public TYPE_INFO ParseCheckHeader()
		{
			GetNext();

			if ( Current_Token != TOKEN.TOK_OPAREN  )
			{
				CSyntaxErrorLog.AddLine("Opening Parenthesis expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"( Parenthesis expected",-1);
                 
			}

			GetNext();

			if ( Current_Token != TOKEN.TOK_CELLNAME)
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
			}

			ArrayList arr = new ArrayList(); 
			String SheetName;

			SheetName = base.cell_value.SheetName;
 
			while ( Current_Token == TOKEN.TOK_CELLNAME )  
			{
				arr.Add(base.cell_value.CellName); 
				// Extracted Value 
				GetNext();
				if ( Current_Token == TOKEN.TOK_COMMA)
				{
					GetNext(); 
				}
				
			}

			if ( Current_Token != TOKEN.TOK_CPAREN )
			{
				CSyntaxErrorLog.AddLine("Invalid Spreadsheet cell reference");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-200," Bad spreadsheet reference",-1);
			}

			         
			CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");

			bool  ret = st.CheckHeader(SheetName,arr);

			ValueStack.push((long)((ret)?1:0)); 			
			GetNext();
			return TYPE_INFO.TYPE_BOOL;  
		}
		
		/// <summary>
		///   Variable declaration
		///   
		/// </summary>
		void VariableStatement(TYPE_INFO type)
		{
			GetNext();
			while ( Current_Token == TOKEN.TOK_UNQUOTED_STRING)
			{
				if ( sym.Lookup(base.last_str) !=null)
				{
					CSyntaxErrorLog.AddLine("Re-defenition of variable"+base.last_str);     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Duplicate Variable",this.SaveIndex());
				}
 
				Symbol rs = new Symbol();
				rs.SymbolName = base.last_str;
				rs.Type = type;
				rs.dbl_val = 0.0;
				sym.Add(base.last_str,rs);
				GetNext();
				if ( Current_Token == TOKEN.TOK_SEMI )
					break;
				else if ( Current_Token == TOKEN.TOK_COMMA )
				{
					GetNext();
					continue;

				}
				else 
				{
					CSyntaxErrorLog.AddLine(", or ; expected");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,", or ; expected",SaveIndex());
				}
			}

			if ( Current_Token != TOKEN.TOK_SEMI )
			{

				CSyntaxErrorLog.AddLine(" ; expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"; expected",SaveIndex());
				

			}
		}
		

		/// <summary>
		///    IF statement
		/// </summary>
		void IfStatement()
		{
			GetNext();
			ValueStack.Clear(); 
			TYPE_INFO RetValue = BExpr();  // Evaluate Expression

			if ( RetValue != TYPE_INFO.TYPE_BOOL )
			{
				
				CSyntaxErrorLog.AddLine(" Boolean Expression expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Boolean Expression expected",SaveIndex());

				
			}
 
			if ( Current_Token != TOKEN.TOK_THEN )
			{
				CSyntaxErrorLog.AddLine(" Then Expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Then Expected",SaveIndex());
				
			}
			////////////////////////////////////////////
			// Save the current state of the lexical analyzer
			//
			//
			int current_index = base.SaveIndex(); 

			GetNext();
			long ds = ValueStack.pop();

			if ( ds == 1 ) // if true 
			{
				StatementList();
				if ( ( Current_Token != TOKEN.TOK_ENDIF ) &&
					( Current_Token != TOKEN.TOK_ELSE ))
				{
					CSyntaxErrorLog.AddLine("End if or Else Expected");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"End if or else Expected",SaveIndex());

				}
				GetNext();
				JumpAfterENDIF(); 
		
			}
			else  // if condition is false
			{
				SkipToProperELSEORENDIF();
				if ( Current_Token != TOKEN.TOK_ENDIF )
				{
					GetNext();
					StatementList();
					if ( ( Current_Token != TOKEN.TOK_ENDIF ))
					{
						CSyntaxErrorLog.AddLine("End if Expected");     
						CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
						throw new CParserException(-100,"End if Expected",SaveIndex());


					}

				}
	

			}
                 
		
		}    
		/// <summary>
		///   While statement
		/// </summary>
		void WhileStatement()
		{
            
			//////////////////////////////////////////////////
			///
			/// Save the state of the lexical analyzer
			///
			///
			int current_index = base.SaveIndex();

			StartLoop:
			base.RestoreIndex(current_index); 
			GetNext();

			TYPE_INFO RetValue = BExpr();

			if ( RetValue != TYPE_INFO.TYPE_BOOL )
			{
				CSyntaxErrorLog.AddLine("Boolean Value Expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Boolean Value Expected",SaveIndex());
				
			}
			long ds = ValueStack.pop();
			if ( ds == 1 ) 
			{
				StatementList();
				if ( ( Current_Token != TOKEN.TOK_WEND ) )
				{
					CSyntaxErrorLog.AddLine("Wend Expected");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Wend Expected",SaveIndex());

				}
				GetNext();
				goto StartLoop;  
			}
			else 
			{
				base.SkipToProperWend();
			}
			
		}
		/// <summary>
		/// 
		/// </summary>
		void AssignmentStatemant()
		{
			String Variable = base.last_str;
   
			Symbol s = (Symbol)sym.Lookup(Variable); 

			
			if ( s == null )
			{
				CSyntaxErrorLog.AddLine("Variable not found  "+last_str);     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Variable not found",SaveIndex());
				
			}

			GetNext();

			if ( Current_Token != TOKEN.TOK_ASSIGN )
			{
				
				CSyntaxErrorLog.AddLine("= expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"= expected",SaveIndex());
				
			}

			GetNext();
 
			TYPE_INFO RetValue = BExpr();

			if ( RetValue != s.Type )
			{
				CSyntaxErrorLog.AddLine("Type mismatch while assigning ");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Type mismatch while assigning",SaveIndex());

				
			}

			if ( RetValue == TYPE_INFO.TYPE_STRING ) 
			{
				long str_index = ValueStack.pop();
				s.str_val =(String) StringList[(int)str_index];
			}
			else if ( RetValue == TYPE_INFO.TYPE_BOOL )
			{
				long bVal = ValueStack.pop();
				s.bol_val = bVal == 0 ? false : true;

			}   
			else  if ( RetValue == TYPE_INFO.TYPE_DOUBLE )
			{
				double rd = ValueStack.PopD();
				s.dbl_val = rd;  
			}
            

			if ( Current_Token != TOKEN.TOK_SEMI )
			{
				CSyntaxErrorLog.AddLine("; expected");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100," ; expected",-1);
				
			}

			
		}
		/// <summary>
		///   
		/// </summary>
		/// <returns></returns>
		public object CallExpr()
		{ 
			
			try 
			{
				ValueStack.Clear();
				TYPE_INFO RetValue = BExpr(); 
				if ( RetValue == TYPE_INFO.TYPE_BOOL )
				{
					long  nd = ValueStack.pop();
					return ( nd == 0 ) ? false : true;
				}
				else if ( RetValue == TYPE_INFO.TYPE_DOUBLE )
				{
					double nd = ValueStack.PopD();  
					return nd;
				}
				else if ( RetValue == TYPE_INFO.TYPE_STRING )
				{
					int nd = (int)ValueStack.pop();
					return StringList[nd].ToString();
                
				}
				else
				{
					CSyntaxErrorLog.AddLine("Type mismatch");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Type mismatch",-1);
					
				}
			}
			catch(CParserException e )
			{
				throw e; 
				
			}


		}
      
		/// <summary>
		///   Support for Boolean Operators
		/// </summary>
		public TYPE_INFO BExpr()
		{

			TYPE_INFO RetValue;
			TOKEN l_token;
			RetValue = LExpr();
			while  ( Current_Token == TOKEN.TOK_AND  || 
				Current_Token == TOKEN.TOK_OR ) 
			{
				TYPE_INFO Temp;
				l_token = Current_Token;
				Current_Token = GetNext();
				Temp = LExpr();

				if ( ( Temp != TYPE_INFO.TYPE_BOOL ) ||
					( RetValue != TYPE_INFO.TYPE_BOOL ))
				{

					CSyntaxErrorLog.AddLine("Boolean Value Expected");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Boolean Value Expected",-1);

					
				}
 
				long x = ValueStack.pop();
				long y = ValueStack.pop();

				switch(l_token )
				{
					case TOKEN.TOK_OR :
						ValueStack.push((long)(x==1 || y == 1  ? 1 : 0 )); 
						break;
					case TOKEN.TOK_AND :
						ValueStack.push((long)(x == 1 &&  y == 1 ? 1 : 0 )); 
						break;
				}
			   
				
			}

			return RetValue;


		}
		/// <summary>
		///   Support for Relational Operators
		/// </summary>
		public TYPE_INFO LExpr()
		{
			TYPE_INFO RetValue;
			TOKEN l_token;
			RetValue = Expr();
			while ( Current_Token == TOKEN.TOK_EQ  || 
				Current_Token == TOKEN.TOK_NEQ ||
				Current_Token == TOKEN.TOK_GT  ||
				Current_Token == TOKEN.TOK_LT  ||
				Current_Token == TOKEN.TOK_LTE ||
				Current_Token == TOKEN.TOK_GTE ) 
			{
				TYPE_INFO Temp;

				l_token = Current_Token;
				Current_Token = GetNext();
				Temp = Expr();

				if ( Temp != RetValue )
				{
					CSyntaxErrorLog.AddLine("Type mismatch");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Type mismatch",-1);
					
				}
 
				if ( Temp == TYPE_INFO.TYPE_DOUBLE )
				{
					double   y = ValueStack.PopD();
					double   x = ValueStack.PopD();
					switch(l_token )
					{
						case TOKEN.TOK_EQ:
							ValueStack.push((long)( x == y  ? 1 : 0)); 
							break;
						case TOKEN.TOK_NEQ :
							ValueStack.push((long)(x != y  ? 1 : 0)); 
							break;
						case TOKEN.TOK_GT:
							ValueStack.push((long)((x > y ) ? 1 : 0)); 
							break;
						case TOKEN.TOK_GTE:
							ValueStack.push((long)(x >= y  ? 1 : 0)); 
							break;
						case TOKEN.TOK_LT:
							ValueStack.push((long)(x < y  ? 1 : 0 )); 
							break;
						case TOKEN.TOK_LTE:
							ValueStack.push((long)(x <= y  ? 1 : 0)); 
							break;
					}    
					RetValue = TYPE_INFO.TYPE_BOOL;

				}
				else if ( Temp == TYPE_INFO.TYPE_BOOL )
				{
					long    y = ValueStack.pop();
					long    x = ValueStack.pop();

					if ( ( x > 1 || y > 1 ) || ( x<0 || y < 0 ))
					{
						
						CSyntaxErrorLog.AddLine("Boolean Value expected");     
						CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
						throw new CParserException(-100,"Boolean Value expected",-1);

					}
						   
					switch(l_token )
					{
						case TOKEN.TOK_EQ:
							ValueStack.push((long)( x == y  ? 1 : 0)); 
							break;
						case TOKEN.TOK_NEQ :
							ValueStack.push((long)(x != y  ? 1 : 0)); 
							break;
						case TOKEN.TOK_GT:
							ValueStack.push((long)((x > y ) ? 1 : 0)); 
							break;
						case TOKEN.TOK_GTE:
							ValueStack.push((long)(x >= y  ? 1 : 0)); 
							break;
						case TOKEN.TOK_LT:
							ValueStack.push((long)(x < y  ? 1 : 0 )); 
							break;
						case TOKEN.TOK_LTE:
							ValueStack.push((long)(x <= y  ? 1 : 0)); 
							break;
					}    
					RetValue = TYPE_INFO.TYPE_BOOL;
				}
				else if ( Temp == TYPE_INFO.TYPE_STRING)
				{
					long    y = ValueStack.pop();
					long    x = ValueStack.pop();

					String first  =(String) StringList[(int)x];
					String second = (String)StringList[(int)y];
						   
					int comp = first.CompareTo(second); 
					switch(l_token )
					{
						case TOKEN.TOK_EQ:
							ValueStack.push((long)( first == second  ? 1 : 0)); 
							break;
						case TOKEN.TOK_NEQ :
							ValueStack.push((long)(first != second  ? 1 : 0)); 
							break;
						case TOKEN.TOK_GT:
							ValueStack.push((long)((comp > 0 ) ? 1 : 0)); 
							break;
						case TOKEN.TOK_GTE:
							ValueStack.push((long)(comp >= 0  ? 1 : 0)); 
							break;
						case TOKEN.TOK_LT:
							ValueStack.push((long)(comp < 0  ? 1 : 0 )); 
							break;
						case TOKEN.TOK_LTE:
							ValueStack.push((long)(comp <= 0  ? 1 : 0)); 
							break;
					}    
					RetValue = TYPE_INFO.TYPE_BOOL;
				}
			}
			return RetValue ;
		}
		/// <summary>
		///    
		/// </summary>


		public TYPE_INFO Expr()
		{
			TYPE_INFO  RetValue;
			TOKEN l_token;
			RetValue = Term();
			while  ( Current_Token == TOKEN.TOK_PLUS  || Current_Token == TOKEN.TOK_SUB ) 
			{
				TYPE_INFO Temp;
				l_token = Current_Token;
				Current_Token = GetNext();
				Temp = Term();

				if ( Temp != RetValue )
				{
					CSyntaxErrorLog.AddLine("Type mismatch");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Type mismatch",-1);
				}

				if ( Temp == TYPE_INFO.TYPE_DOUBLE ) 
				{
					double x = ValueStack.PopD();
					double y = ValueStack.PopD();
					ValueStack.push( (l_token == TOKEN.TOK_PLUS ) ? (x + y) : (y-x) ); 
				}
				else if ( Temp == TYPE_INFO.TYPE_STRING )
				{
                    

					long s = ValueStack.pop();
					long d = ValueStack.pop();

					String first  = (String)StringList[(int)s];
					String second = (String)StringList[(int)d];
					long ru;

					if ( l_token == TOKEN.TOK_PLUS )
						ru = StringList.Add(second + first); 
					else
						ru = StringList.Add(second.TrimEnd()+first.TrimStart());
     
					ValueStack.push((long)ru);
                      
				}
			}
  
			return RetValue;

		}
		/// <summary>
		/// 
		/// </summary>

		public TYPE_INFO Term()
		{
			TOKEN l_token;
			TYPE_INFO RetValue = TYPE_INFO.TYPE_ILLEGAL;
 
			RetValue = Factor();

			while ( Current_Token == TOKEN.TOK_MUL  || Current_Token == TOKEN.TOK_DIV ) 
			{
				TYPE_INFO Temp = TYPE_INFO.TYPE_ILLEGAL; 
				l_token = Current_Token;
				Current_Token = GetNext();
				Temp = Factor();

				if ( ( RetValue != Temp ) )
				{
					CSyntaxErrorLog.AddLine("Type mismatch");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Type mismatch",-1);
				}

 

				double x = ValueStack.PopD();
				double y = ValueStack.PopD();

				if ( x == 0 ) 
				{ 
					CSyntaxErrorLog.AddLine("Division by Zero error");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Division by Zero",-1);
					
				
				}
				ValueStack.push( (l_token == TOKEN.TOK_MUL ) ? (x * y) :  (y/x) ); 
			}
			return RetValue;

		}

		public TYPE_INFO Factor()
		{
			TOKEN l_token;

			TYPE_INFO RetValue = TYPE_INFO.TYPE_ILLEGAL;
 
			if ( Current_Token == TOKEN.TOK_DOUBLE )
			{

				ValueStack.push(GetNumber());
				Current_Token = GetNext();  
				RetValue = TYPE_INFO.TYPE_DOUBLE; 
			} 
			else if ( Current_Token == TOKEN.TOK_OPAREN )
			{

				Current_Token = GetNext();   

				RetValue = BExpr();  // Recurse

				if ( Current_Token != TOKEN.TOK_CPAREN )
				{
					CSyntaxErrorLog.AddLine("Missing ) ");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Missing )",-1);
					
					
				}   
				Current_Token = GetNext();            
			} 

			else if ( Current_Token == TOKEN.TOK_PLUS 
				|| Current_Token == TOKEN.TOK_SUB
				|| Current_Token == TOKEN.TOK_NOT )
			{
				l_token = Current_Token;
				Current_Token = GetNext();
				RetValue = Factor();

				if ( l_token != TOKEN.TOK_NOT) 
				{
					double  x = ValueStack.PopD();
					if ( l_token == TOKEN.TOK_SUB )
						x = -x;
					ValueStack.push(x);  
				}
				else 
				{
					long xt = ValueStack.pop();
					if ( !( xt == 0 || xt == 1 ) )
					{
						CSyntaxErrorLog.AddLine("Boolean value expected");     
						CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
						throw new CParserException(-100,"Boolean value expected",-1); 
						
					}
					ValueStack.push((long)(xt == 0 ? 1 : 0 )); 
				}
   
			}

			else if (Current_Token == TOKEN.TOK_BOOL_FALSE  ||
				Current_Token == TOKEN.TOK_BOOL_TRUE  )
			{
				ValueStack.push((long)((Current_Token == TOKEN.TOK_BOOL_FALSE )?0:1)); 
				RetValue = TYPE_INFO.TYPE_BOOL;  
				GetNext();
			}
			else if ( Current_Token == TOKEN.TOK_CELL)
			{
				
				CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");
				String str = st.GetCellValue(base.cell_value.SheetName ,
					base.cell_value.CellName );  
				int rst = StringList.Add(str);
				ValueStack.push((long)rst); 
				Current_Token = GetNext();  
				RetValue = TYPE_INFO.TYPE_STRING;  

			}
			else if ( Current_Token == TOKEN.TOK_CELLDBCOLUMN )
			{
				String str_name   = base.cell_tbl_clm.SheetName ;
				String str_column = base.cell_tbl_clm.CellName ;
				GetNext();
				RetValue = BExpr();

				if ( RetValue != TYPE_INFO.TYPE_DOUBLE )
				{
					CSyntaxErrorLog.AddLine("Number expected");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Number expected",-1); 
					
				}

				double temp  = ValueStack.PopD();
				base.cell_tbl_clm.Index = ( int)temp; 
				/////////////////////////////////
				/// Remove Later
				///
				///

				CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");
				String srt = st.GetColumnValue(base.cell_tbl_clm.SheetName ,
					base.cell_tbl_clm.CellName ,
					base.cell_tbl_clm.Index ); 
				int rs = StringList.Add(srt);
				ValueStack.push((long)rs); 

				Current_Token = GetNext();

				RetValue = TYPE_INFO.TYPE_STRING;
				
			}
			else if ( Current_Token == TOKEN.TOK_DBCOLUMN )
			{
				ValueStack.push(1.0); // Replace the Code here 
				Current_Token = GetNext();  
				RetValue = TYPE_INFO.TYPE_DOUBLE;  
             
			}
			else if ( Current_Token == TOKEN.TOK_BUILTIN_SUM_RANGE )
				
			{
				
				RetValue = ParseSum();
			
			
			}
			else if ( Current_Token == TOKEN.TOK_BUILTIN_LOOKUP )
			{
				RetValue = ParseLookup(true);
				 

			}
			else if ( Current_Token == TOKEN.TOK_BUILTIN_LOOKUP_CASE_IGNORE )
			{
				RetValue = ParseLookup(false);

			}
			else if ( Current_Token == TOKEN.TOK_CHECKDUP )
			{
				RetValue = ParseCheckDup();

			}
			else if ( Current_Token == TOKEN.TOK_CHECKNULL )
			{
                
               RetValue = this.ParseCheckNull(); 
			}
			else if ( Current_Token == TOKEN.TOK_CLEANUP )
			{
				RetValue = ParseCleanup(); 

			}
			else if ( Current_Token == TOKEN.TOK_CHECKCURRENCY )
			{
				RetValue = ParseIsCurrency(); 

			}
			else if ( Current_Token == TOKEN.TOK_CHECKHEADER )
			{
				RetValue = ParseCheckHeader(); 

			}
			else if ( Current_Token == TOKEN.TOK_CHECK_HOUR )
			{
				RetValue = ParseCheckHourly();

			}
			else if ( Current_Token ==  TOKEN.TOK_RECORD_COUNT )
			{
				RetValue = ParseRecCount();  

			}
			else if ( Current_Token == TOKEN.TOK_UNQUOTED_STRING )
			{
				Symbol st =(Symbol)sym.Lookup(base.last_str);
  
				if ( st == null )
				{
					CSyntaxErrorLog.AddLine("Variable expected");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Variable Expected",-1); 
					

				}

				if ( st.Type == TYPE_INFO.TYPE_BOOL )
				{
					ValueStack.push((long) (st.bol_val ? 1 : 0));  
					
				}
				else if ( st.Type  == TYPE_INFO.TYPE_DOUBLE )
				{
					ValueStack.push(st.dbl_val);
                    
				}
				else 
				{
					int rst = StringList.Add(st.str_val);
					ValueStack.push((long)rst);   

				}
				GetNext();


				RetValue = st.Type;

			}
			else if ( Current_Token == TOKEN.TOK_BUILTIN_VAL )
			{
				RetValue = ParseVal();   
			} 
			else if ( Current_Token ==	TOKEN.TOK_SCAN_BRANDID )
			{
				RetValue = ParseScanBrand();
			}
			else if ( Current_Token == TOKEN.TOK_TRIM  ||
				Current_Token == TOKEN.TOK_UPPER )
			{
				RetValue = (Current_Token == TOKEN.TOK_TRIM )?ParseTrim() : ParseUpper();
                 
			}
			else if ( Current_Token == TOKEN.TOK_SCAN_PULL )
			{

				RetValue = ScanPullStatement();   
			}
			else if ( Current_Token == TOKEN.TOK_SCAN )
			{

				RetValue = ScanStatement();   
			}
			else if ( Current_Token == TOKEN.TOK_STRING )
			{
				int index = StringList.Add(base.last_str);
				ValueStack.push(index); 
				RetValue = TYPE_INFO.TYPE_STRING; 
				GetNext();
  
			}
			else 
			{
				CSyntaxErrorLog.AddLine("Illegal Token in input");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"Illegal Token",-1); 

				
			} 

			return RetValue;

		}
		/// <summary>
		///    Ctor
		/// </summary>
		/// <param name="str"></param>
		public CParser(String str,String Country,String Segment ):base(str)
		{
			CountryFromFolder = Country;
			SegmentFromFolder = Segment;

		}
		/// <summary>
		/// 
		/// </summary>

		public void CloseAll()
		{
            if ( sym == null )
				   return;

            CExcelReader st = ( CExcelReader ) sym.Lookup("EXCEL_READER");

			if ( st == null )
			{
               System.Diagnostics.EventLog.WriteEntry("ForeCastLog","Diagnostic message := st == null");  
			}
			else 
			{
				st.Close();
			}

			
		
			 
		}
		/// <summary>
		///    CallParser
		///    Entry point to the Parser
		///    Pass the Excel File name as the parameter
		/// </summary>

		public bool CallParser(String ExcelConStr , String OracleConStr)
		{
			try 
			{
				OpenStatement(ExcelConStr,OracleConStr);  
				GetNext();   
				StatementList();
				CloseAll();
				return false;
			}
			catch(CParserException e)
			{
				if ( e.GetErrorCode() == 200 )
				{
					CSemanticErrorLog.AddLine("Successful Validation"); 
					
					return true;
				}
				else 
				{
					System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString());
					CSyntaxErrorLog.AddLine(e.ToString());    
					CloseAll();
					return false;
				}
			}
			catch(Exception e )
			{

				System.Diagnostics.EventLog.WriteEntry("ForeCastLog",e.ToString());
				CSyntaxErrorLog.AddLine(e.ToString());    
				CloseAll();
				return false;

			}
			
			

		}
		/// <summary>
		/// 
		/// </summary>

		public void Statement()
		{

			switch(Current_Token)
			{
				case TOKEN.TOK_BEGIN:
					BeginStatement();
					GetNext();
					break;
				case TOKEN.TOK_VALIDATE:
					ValidateStatement();
					GetNext();
					break; 
				case TOKEN.TOK_VAR_STRING:
					VariableStatement(TYPE_INFO.TYPE_STRING);
					GetNext();
					break;
				case TOKEN.TOK_VAR_NUMBER:
					VariableStatement(TYPE_INFO.TYPE_DOUBLE);
					GetNext();
					break;
				case TOKEN.TOK_VAR_BOOL:
					VariableStatement(TYPE_INFO.TYPE_BOOL);
					GetNext();
					break;
				case TOKEN.TOK_IF:
					IfStatement();
					GetNext();
					break;
				case TOKEN.TOK_WHILE:
					WhileStatement();
					GetNext();
					break;
				case TOKEN.TOK_UNQUOTED_STRING:
					AssignmentStatemant();
					GetNext();
					break;
				case TOKEN.TOK_WRITELOG:
					WriteLogStatement();
					GetNext();
					break;
				default:
				{
					CSyntaxErrorLog.AddLine("Invalid Statement");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"Invalid Statement",-1); 

                    
				}
					

			}

			

		}
		/// <summary>
		/// 
		/// </summary>

		public void StatementList()
		{
			/////////////////////////////////
			///
			///   
			///
			while ( ( Current_Token != TOKEN.TOK_STOP ) &&
				    ( Current_Token != TOKEN.TOK_ELSE ) &&
				    ( Current_Token != TOKEN.TOK_ENDIF ) &&
				    ( Current_Token != TOKEN.TOK_WEND ) &&
				    ( Current_Token != TOKEN.TOK_END  ) &&
				    ( Current_Token != TOKEN.TOK_EXCEPTION)
				)
			{
				Statement();	
			}


			///
			///  if Stop Statement 
			///
			///
			if ( Current_Token == TOKEN.TOK_STOP )
			{
                    StopStatement();  
			}

		}


		


	}

	//////////////////////////////////////////////////////////
	//
	//
	//
	//
	public class Lexer
	{
		
		/// <summary>
		///    Items which are static of nature
		/// </summary>
		String IExpr;            // Expression string
		protected ValueTable  [] keyword; // Keyword Table
		int    length;           // Length of the string 
		double number;           // Last grabbed number from the stream

		/// <summary>
		///  Items which are dependent on state
		///  index can be changed by GetNext,a Loop or IF statement
		/// </summary>
		int    index ;           // index into a character  

		/// <summary>
		///        last_str := Token assoicated with 
		/// </summary>

		public String last_str; // Last grabbed String
		public CellValue cell_value;  // Cell Value
		public CTableColumn tbl_clm;   // DB reference
		public CellAsTableColumn cell_tbl_clm; //Cell as Column

		/// <summary>
		///    Current Token and Last Grabbed Token
		/// </summary>
		protected TOKEN Current_Token;  // Current Token
		protected TOKEN Last_Token;     // Penultimate token


		

		/// <summary>
		///    Get Next Token from the stream and return to the caller
		/// </summary>
		/// <returns></returns>

		protected TOKEN GetNext()
		{
			Last_Token = Current_Token;
			Current_Token = GetToken(); 
			return Current_Token;
		}
		/// <summary>
		///   Save the Current Lexer index
		/// </summary>
		/// <returns></returns>
		public int SaveIndex()
		{
			return index;
		}
		/// <summary>
		///     Get Line Correswhere Error Occured
		/// </summary>
		/// <param name="pindex"></param>
		/// <returns></returns>
		public String GetCurrentLine(int pindex)
		{
			int tindex = pindex;
			if ( pindex >= length )
			{
                   tindex = length-1;
			}
			while ( tindex > 0  && IExpr[tindex] != '\n' )
                   tindex--;

			if ( IExpr[tindex] == '\n' )
				    tindex++;

            String CurrentLine = "";

			while (tindex <length &&  ( IExpr[tindex] !='\n' ) )
			{
               CurrentLine = CurrentLine + IExpr[tindex];
               tindex++; 
			}

			return CurrentLine+"\n";

		}
		/// <summary>
		///    
		/// </summary>
		/// <param name="index"></param>
		/// <returns></returns>

		public String GetPreviousLine( int pindex )
		{

			int tindex = pindex;
			while ( tindex > 0  && IExpr[tindex] != '\n' )
				          tindex--;

			if ( IExpr[tindex] == '\n' )
				tindex--;
            else
				return "";

			while ( tindex > 0  && IExpr[tindex] != '\n' )
				tindex--;


			if ( IExpr[tindex] == '\n' )
				tindex--;


			String CurrentLine = "";

			while (tindex <length &&  ( IExpr[tindex] !='\n' ) )
			{
				CurrentLine = CurrentLine + IExpr[tindex];
				tindex++; 
			}

			return CurrentLine+"\n";



		}
		/// <summary>
		///    Restore Index . Only after a GetNext , contents
		///    of the state variables will be reliable
		/// </summary>
		/// <param name="m_index"></param>
		public void RestoreIndex(int m_index)
		{
			index = m_index;

		}
		/// <summary>
		///    Skip to the end statement
		/// </summary>

		public void SkipToEnd()
		{

			long count_if = 0;    // Count # of IF
			long count_while=0;   // Count # of WHILE
			long count_begin=0;   // Count # of BEGIN block
			long count_recover=0; // Count # of Recover block
			long count_else=0;    // Count # of ELSE
			
			/////////////////////////////////
			///
			/// Nesting stack will keep track of if/while/begin nesting
			///

			Stack nesting_stack = new Stack(); 
			Stack else_stack    = new Stack();  // Stack for tracking ELSE mismatch
			Stack recover_stack = new Stack();
 

			////////////////////////
			///
			/// Label to act as a psuedo loop
			///
			skip_eol:

				/////////////////////////////////////////////////////
				/// if any of them become less than zero fatal error
				///
				if ( count_if    < 0 || 
					count_while < 0 || 
					count_begin < 0 || 
					count_recover < 0 ||
					count_else    < 0 

					)
				{
					
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
					
				}

			if ( Current_Token == TOKEN.TOK_WHILE)
			{
				nesting_stack.push(1); 
				count_while++;
			}
			else if ( Current_Token == TOKEN.TOK_IF )
			{
				nesting_stack.push(0);
				else_stack.push(count_else);   
				count_if++;
				count_else++;
			}
			else if ( Current_Token == TOKEN.TOK_BEGIN )
			{
				nesting_stack.push(2);
				recover_stack.push(count_recover);  
				count_begin++;
				count_recover++;
			}
			else if  ( ( Current_Token == TOKEN.TOK_ENDIF ) ||
				( Current_Token == TOKEN.TOK_ELSE  ) )
			{
							
				long x = nesting_stack.pop();
				if ( x != 0 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
					
				}

				if ( Current_Token == TOKEN.TOK_ENDIF )
				{
					count_if--;
					long rst = else_stack.pop();
  
					if ( rst != count_if )
					{
						CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
						CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
						throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
						
					}

				}
				else 
				{
					if ( else_stack.IsEmpty() )
					{
						CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
						CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
						throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
					}


					long r = else_stack.pop();
					count_else--;

					if ( r != count_else )
					{
						CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
						CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
						throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
						
					}
					else 
					{
						nesting_stack.push(x); 
						else_stack.push(r); 
					}
				      
				} 
			}
			else if ( Current_Token == TOKEN.TOK_WEND )
			{
				long x = nesting_stack.pop();
 
				if ( x != 1 )
				{
			    	CSyntaxErrorLog.AddLine("While/Wend nesting error");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"While/Wend nesting error",-1); 
					

				}
				count_while--;
				
			}
			else if ( Current_Token == TOKEN.TOK_END )
			{
				if ( count_if    == 0 &&  count_while ==  0 && 
					count_begin == 0 &&  count_recover == 0  
					) 
				{
					return;
				}

				long x = nesting_stack.pop();

				if ( x != 2 )
				{
					CSyntaxErrorLog.AddLine("BEGIN...EXCEPTION..END nesting error");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"BEGIN...EXCEPTION..END nesting error",-1); 
					

				}
				count_begin--;

			}
			else if ( Current_Token == TOKEN.TOK_EXCEPTION )
			{
				long x = nesting_stack.pop();
				if ( x != 2 )
				{
					CSyntaxErrorLog.AddLine("BEGIN...EXCEPTION..END nesting error");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"BEGIN...EXCEPTION..END nesting error",-1); 
					

				}
				nesting_stack.push(x);

				long r = recover_stack.pop();
				count_recover--;
				if ( r != count_recover )
				{
					CSyntaxErrorLog.AddLine("BEGIN...EXCEPTION..END nesting error");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"BEGIN...EXCEPTION..END nesting error",-1); 
				}
				else 
				{
					recover_stack.push(r); 
				}
				               
				
			}


			if ( Current_Token == TOKEN.ILLEGAL_TOKEN ||
				Current_Token == TOKEN.TOK_NULL )
			{
				CSyntaxErrorLog.AddLine("BEGIN...EXCEPTION..END nesting error");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"BEGIN...EXCEPTION..END nesting error",-1); 
			
			}

			SkipToEoln();
			GetNext();
			goto skip_eol;

		}

		/// <summary>
		///    Skip to the end statement
		/// </summary>

		public void SkipToException()
		{
			long count_if = 0;    // Count # of IF
			long count_while=0;   // Count # of WHILE
			long count_begin=0;   // Count # of BEGIN block
			long count_recover=0; // Count # of Recover block
			long count_else=0;    // Count # of ELSE
			
			/////////////////////////////////
			///
			/// Nesting stack will keep track of if/while/begin nesting
			///

			Stack nesting_stack = new Stack(); 
			Stack else_stack    = new Stack();  // Stack for tracking ELSE mismatch
			Stack recover_stack = new Stack();
 

			////////////////////////
			///
			/// Label to act as a psuedo loop
			///
			skip_eol:

				/////////////////////////////////////////////////////
				/// if any of them become less than zero fatal error
				///
				if ( count_if    < 0 || 
					count_while < 0 || 
					count_begin < 0 || 
					count_recover < 0 ||
					count_else    < 0 

					)
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
					
				}

			if ( Current_Token == TOKEN.TOK_WHILE)
			{
				nesting_stack.push(1); 
				count_while++;
			}
			else if ( Current_Token == TOKEN.TOK_IF )
			{
				nesting_stack.push(0);
				else_stack.push(count_else);   
				count_if++;
				count_else++;
			}
			else if ( Current_Token == TOKEN.TOK_BEGIN )
			{
				nesting_stack.push(2);
				recover_stack.push(count_recover);  
				count_begin++;
				count_recover++;
			}
			else if  ( ( Current_Token == TOKEN.TOK_ENDIF ) ||
				( Current_Token == TOKEN.TOK_ELSE  ) )
			{
							
				long x = nesting_stack.pop();
				if ( x != 0 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
					
				}

				if ( Current_Token == TOKEN.TOK_ENDIF )
				{
					count_if--;
					long rst = else_stack.pop();
  
					if ( rst != count_if )
					{
						CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
						CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
						throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
						
					}

				}
				else 
				{
					if ( else_stack.IsEmpty() )
					{
						CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
						CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
						throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
						
					}


					long r = else_stack.pop();
					count_else--;

					if ( r != count_else )
					{
						CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
						CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
						throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
					}
					else 
					{
						nesting_stack.push(x); 
						else_stack.push(r); 
					}
				      
				} 
			}
			else if ( Current_Token == TOKEN.TOK_WEND )
			{
				long x = nesting_stack.pop();
 
				if ( x != 1 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 

				}
				count_while--;
				
			}
			else if ( Current_Token == TOKEN.TOK_END )
			{
				if ( count_if    == 0 &&  count_while ==  0 && 
					count_begin == 0 &&  count_recover == 0  
					) 
				{
					return;
				}

				long x = nesting_stack.pop();

				if ( x != 2 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 



				}
				count_begin--;

			}
			else if ( Current_Token == TOKEN.TOK_EXCEPTION )
			{

				if ( count_if    == 0 &&  count_while ==  0 && 
					count_begin == 0 &&  count_recover == 0  
					) 
				{
					return;
				}
				long x = nesting_stack.pop();
				if ( x != 2 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 


				}
				nesting_stack.push(x);

				long r = recover_stack.pop();
				count_recover--;
				if ( r != count_recover )
				{

					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
				}
				else 
				{
					recover_stack.push(r); 
				}
				               
				
			}


			if ( Current_Token == TOKEN.ILLEGAL_TOKEN ||
				Current_Token == TOKEN.TOK_NULL )
			{
				CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
				
				 

			}

			SkipToEoln();
			GetNext();
			goto skip_eol;
  



		}
		/// <summary>
		/// Jump to the statement after the endif
		/// </summary>

		public void JumpAfterENDIF()
		{
			int count_if = 0;    // Count # of IF
			int count_while=0;   // Count # of WHILE
			int count_begin=0;   // Count # of BEGIN block
			int recover_count=0; // Count # of Recover block

			Stack while_stack = new Stack(); 

			skip_eol:

				if ( count_if < 0 || count_while < 0 || count_begin < 0 || recover_count < 0 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
					
				}

			if ( Current_Token == TOKEN.TOK_WHILE)
			{
				while_stack.push(1);
				count_while++;
			}
			else if ( Current_Token == TOKEN.TOK_IF )
			{
				while_stack.push(0);
				count_if++;
			
			}
			else if ( Current_Token == TOKEN.TOK_BEGIN )
			{
				while_stack.push(2);
				count_begin++;
				recover_count++;

			}
			else if  ( ( Current_Token == TOKEN.TOK_ENDIF ) ||
				( Current_Token == TOKEN.TOK_ELSE  ) )
			{
				
							
				////////////////////////
				///
				///  Return if proper ENDIF found
				///
				if ( count_if == 0 
					&& count_while == 0 
					&& count_begin == 0 && recover_count == 0 )
					return;

				long x = while_stack.pop();
 
				if ( x != 0 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
					

				}
				count_if--;

				
			
			}
			else if ( Current_Token == TOKEN.TOK_WEND )
			{
				long x = while_stack.pop();
 
				if ( x != 1 )
				{
					
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 


				}
				count_while--;
				
			}
			else if ( Current_Token == TOKEN.TOK_END )
			{
				long x = while_stack.pop();
				if ( x != 2 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 

					

				}
				count_begin--;

			}
			else if ( Current_Token == TOKEN.TOK_EXCEPTION )
			{
				long x = while_stack.pop();
				if ( x != 2 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 

				}
				while_stack.push(2);
				recover_count--;

				
			}


			if ( Current_Token == TOKEN.ILLEGAL_TOKEN ||
				Current_Token == TOKEN.TOK_NULL )
			{
				CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
				
				 

			}

			SkipToEoln();
			GetNext();
			goto skip_eol;

		




		}
		/// <summary>
		///    Skip to Proper ELSE OR ENDIFIF
		/// </summary>
		/// <param name="val"></param>

		public void SkipToProperELSEORENDIF()
		{
			int count_if = 0;    // Count # of IF
			int count_while=0;   // Count # of WHILE
			int count_begin=0;   // Count # of BEGIN block
			int recover_count=0; // Count # of Recover block
			////////////////////////////////////////////
			///
			///   Stack to keep track of IF/WHILE nesting 
			///
			Stack while_stack = new Stack();

            
			    
			skip_eol:
                   

				if ( count_if < 0 || count_while < 0 || count_begin < 0 || recover_count < 0 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
					
				}

			if ( Current_Token == TOKEN.TOK_WHILE)
			{
				while_stack.push(1);
				count_while++;
			}
			else if ( Current_Token == TOKEN.TOK_IF )
			{
				while_stack.push(0);
				count_if++;
			
			}
			else if ( Current_Token == TOKEN.TOK_BEGIN )
			{
				while_stack.push(2);
				count_begin++;
				recover_count++;

			}
			else if  ( ( Current_Token == TOKEN.TOK_ENDIF ) ||
				( Current_Token == TOKEN.TOK_ELSE  ) )
			{
				
							
				////////////////////////
				///
				///  Return if proper ENDIF found
				///
				if ( count_if == 0 
					&& count_while == 0 
					&& count_begin == 0 && recover_count == 0 )
					return;

				long x = while_stack.pop();
 
				if ( x != 0 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
					  

				}
				count_if--;

				
			
			}
			else if ( Current_Token == TOKEN.TOK_WEND )
			{
				long x = while_stack.pop();
 
				if ( x != 1 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 

				}
				count_while--;
				
			}
			else if ( Current_Token == TOKEN.TOK_END )
			{
				long x = while_stack.pop();
				if ( x != 2 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 

				}
				count_begin--;

			}
			else if ( Current_Token == TOKEN.TOK_EXCEPTION )
			{
				long x = while_stack.pop();
				if ( x != 2 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 

				}
				while_stack.push(2);
				recover_count--;

				
			}


			if ( Current_Token == TOKEN.ILLEGAL_TOKEN ||
				Current_Token == TOKEN.TOK_NULL )
			{
				CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
				 

			}

			SkipToEoln();
			GetNext();
			goto skip_eol;

		

		}

		/// <summary>
		///     Skip to Proper Wend
		/// </summary>

		public void SkipToProperWend()
		{
			int count_if = 0;    // Count # of IF
			int count_while=0;   // Count # of WHILE
			int count_begin=0;   // Count # of Begin
			int count_recover=0; // Count # of Error
			int count_else=0;    // Count # of else we can expect
			

			////////////////////////////////////////////
			///
			///   Stack to keep track of IF/WHILE nesting 
			///
			Stack while_stack = new Stack();
            
            
			
			    
			skip_eol:
                   
				if (count_else < 0 || count_if < 0 || count_while < 0 || count_begin < 0 || count_recover < 0 )
				{

					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
				}

			if ( Current_Token == TOKEN.TOK_WHILE)
			{
				while_stack.push(1);
				count_while++;
			}
			else if ( Current_Token == TOKEN.TOK_IF )
			{
				while_stack.push(0);
				count_if++;
				count_else++;
			}
			else if ( Current_Token == TOKEN.TOK_BEGIN )
			{
				while_stack.push(0);
				count_begin++;
				count_recover++;

			}
			else if  ( ( Current_Token == TOKEN.TOK_ENDIF ) ||
				( ( Current_Token == TOKEN.TOK_ELSE ) ))
			{
				long x = while_stack.pop();
				if ( x != 0 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 

				}

				if ( Current_Token == TOKEN.TOK_ENDIF ) 
				{
					count_if--;
					
				}
				else 
				{
					while_stack.push(0); 
					count_else--;

				}  
			
			}
			else if ( Current_Token == TOKEN.TOK_WEND )
			{
				
				if ( count_while == 0  && count_if == 0 && 
					count_else  == 0 && count_begin == 0 &&
					count_recover == 0 )
					return;

				long x = while_stack.pop();
 
				if ( x != 1 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 

				}
				
				count_while--;
				
			}
			else if ( Current_Token == TOKEN.TOK_END )
			{

				long x = while_stack.pop();
 
				if ( x != 2 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 

				}
				
				count_begin--;
                  

			}
			else if ( Current_Token == TOKEN.TOK_EXCEPTION )
			{

				long x = while_stack.pop();
 
				if ( x != 2 )
				{
					CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
					CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
					throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 

				}
				while_stack.push(2); 
				count_recover--;
			}



			if ( Current_Token == TOKEN.ILLEGAL_TOKEN ||
				Current_Token == TOKEN.TOK_NULL )
			{
				CSyntaxErrorLog.AddLine("IF/ELSE/BEGIN  NESTING MISMATCH");     
				CSyntaxErrorLog.AddLine(GetCurrentLine(SaveIndex())); 
				throw new CParserException(-100,"IF/ELSE/BEGIN  NESTING MISMATCH",-1); 
				 

			}

			SkipToEoln();
			GetNext();
			goto skip_eol;

		

				    

		}

		/////////////////////////////////////////////
		//
		// Ctor
		//
		//   
		public Lexer(String Expr)
		{
			IExpr = Expr;
			length = IExpr.Length; 
			index = 0;
			////////////////////////////////////////////////////
			// Fill the Keywords
			//
			//
			keyword = new ValueTable [33 ];
			keyword[0] = new ValueTable(TOKEN.TOK_STOP , "STOP" );
			keyword[1] = new ValueTable(TOKEN.TOK_VALIDATE,"VALIDATE");
			keyword[2] = new ValueTable(TOKEN.TOK_BUILTIN_SUM_RANGE,"SUM_RANGE");
			keyword[3] = new ValueTable(TOKEN.TOK_BUILTIN_LOOKUP ,"LOOKUP");
			keyword[4] = new ValueTable(TOKEN.TOK_BOOL_FALSE,"FALSE");
			keyword[5] = new ValueTable(TOKEN.TOK_BOOL_TRUE,"TRUE");  
			keyword[6] = new ValueTable(TOKEN.TOK_BUILTIN_VAL,"VAL");  
			keyword[7] = new ValueTable(TOKEN.TOK_IF,"IF");
			keyword[8]= new ValueTable(TOKEN.TOK_WHILE,"WHILE");
			keyword[9]= new ValueTable(TOKEN.TOK_VAR_STRING,"STRING");
			keyword[10]= new ValueTable(TOKEN.TOK_VAR_BOOL,"BOOLEAN");
			keyword[11]= new ValueTable(TOKEN.TOK_VAR_NUMBER,"NUMERIC");
			keyword[12]= new ValueTable(TOKEN.TOK_WEND,"WEND");
			keyword[13]= new ValueTable(TOKEN.TOK_ELSE,"ELSE");
			keyword[14]= new ValueTable(TOKEN.TOK_ENDIF,"ENDIF"); 
			keyword[15]= new ValueTable(TOKEN.TOK_THEN ,"THEN");
			keyword[16]= new ValueTable(TOKEN.TOK_EXCEPTION,"EXCEPTION");
			keyword[17]= new ValueTable(TOKEN.TOK_BEGIN,"BEGIN");  
			keyword[18]= new ValueTable(TOKEN.TOK_END,"END" );  
			keyword[19]= new ValueTable(TOKEN.TOK_WRITELOG,"WRITELOG");
			keyword[20]= new ValueTable(TOKEN.TOK_SCAN_PULL,"SCAN_PULL");
			keyword[21]= new ValueTable(TOKEN.TOK_SCAN,"SCAN");
			keyword[22]= new ValueTable(TOKEN.TOK_CHECKDUP,"DUPE_CHECK"); 
			keyword[23]= new ValueTable(TOKEN.TOK_CLEANUP,"CLEANUP");
			keyword[24]= new ValueTable(TOKEN.TOK_CHECKHEADER,"CHECKHEADER");
			keyword[25]= new ValueTable(TOKEN.TOK_SCAN_BRANDID,"SCAN_BRANDID");
			keyword[26]= new ValueTable(TOKEN.TOK_CHECKCURRENCY,"ISVALIDCURRENCY");  
			keyword[27]= new ValueTable(TOKEN.TOK_RECORD_COUNT,"RECORDCOUNT");
            keyword[28]= new ValueTable(TOKEN.TOK_UPPER,"UPPER");
			keyword[29]= new ValueTable(TOKEN.TOK_TRIM,"TRIM");
			keyword[30]= new ValueTable(TOKEN.TOK_CHECK_HOUR,"CHECK_HOURLYDISTRIBUTION");
			keyword[31]= new ValueTable(TOKEN.TOK_BUILTIN_LOOKUP_CASE_IGNORE,"LOOKUP_CASE_SENSITIVE");
			keyword[32]= new ValueTable(TOKEN.TOK_CHECKNULL,"CHECK_FOR_NULL");

 

  

  
  
  

  
            			
		}
		/// <summary>
		///      Extract string from the stream
		/// </summary>
		/// <returns></returns>
		private String ExtractString()
		{
			String ret_value = "";
			while ( index < IExpr.Length &&
				(char.IsLetterOrDigit(IExpr[index]) || IExpr[index]=='_' ))
			{
				ret_value = ret_value + IExpr[index];
				index++;
			}
			return ret_value; 
		}

		/// <summary>
		///    Skip to the End of Line
		/// </summary>
		public void SkipToEoln()
		{
			while (index < length && (IExpr[index] != '\r'))
				index++; 

			if ( index == length )
				return;

			if ( IExpr[index+1] == '\n' )
			{
				index+=2;
				return;
			}
			index++;
			return;
		}
		/////////////////////////////////////////////////////
		// Grab the next token from the stream
		//
		//    
		//
		//
		public TOKEN GetToken()
		{
          
			TOKEN tok; 
				
			re_start:
				tok = TOKEN.ILLEGAL_TOKEN;

			////////////////////////////////////////////////////////////
			//
			// Skip  the white space 
			//  

			while (index < length && 
				(IExpr[index] == ' ' || IExpr[index]== '\t') )
				index++;
			//////////////////////////////////////////////
			//
			//   End of string ? return NULL;
			//

			if ( index == length)
				return TOKEN.TOK_NULL;

			/////////////////////////////////////////////////  
			//
			//
			// 
			switch(IExpr[index])
			{
				case '\n':
					index++;
					goto re_start;
				case '\r':
					index++;
					goto re_start;
				case '+':
					tok = TOKEN.TOK_PLUS;
					index++;
					break;
				case '-':
					tok = TOKEN.TOK_SUB;
					index++;
					break;
				case '*':
					tok = TOKEN.TOK_MUL;
					index++;
					break;
				case ',':  
					tok = TOKEN.TOK_COMMA;
					index++;
					break;
				case '(':
					tok = TOKEN.TOK_OPAREN;
					index++;
					break;
				case ';':
					tok = TOKEN.TOK_SEMI;
					index++;
					break;
				case ')':
					tok = TOKEN.TOK_CPAREN;
					index++;
					break;
				case '!':
					tok = TOKEN.TOK_NOT;
					index++;
					break;
				case ':':
					if ( IExpr[index+1] == '=' )
					{
						tok = TOKEN.TOK_COLON;
						index +=2;
					}
					else 
					{
						tok = TOKEN.ILLEGAL_TOKEN;
						index++;
					}
					break;
				case '>':
					if ( IExpr[index+1] == '=' )
					{
						tok = TOKEN.TOK_GTE;
						index +=2;
					}
					else 
					{
						tok = TOKEN.TOK_GT;
						index++;
					}
					break;
				case '<':
					if ( IExpr[index+1] == '=' )
					{
						tok = TOKEN.TOK_LTE;
						index +=2;
					}
					else if ( IExpr[index+1] == '>')
					{
						tok = TOKEN.TOK_NEQ;
						index +=2; 
					}
					else  
					{
						tok = TOKEN.TOK_LT;
						index++;
					}
					break;
				case '=':
					if ( IExpr[index+1] == '=' )
					{
						tok = TOKEN.TOK_EQ;
						index +=2;
					}
					else
					{
						tok = TOKEN.TOK_ASSIGN;
						index++;
					}
					break;
				case '&':
					if ( IExpr[index+1] == '&' )
					{
						tok = TOKEN.TOK_AND;
						index +=2;
					}
					else 
					{
						tok = TOKEN.ILLEGAL_TOKEN;
						index++;
					}
					break;
				case '|':
					if ( IExpr[index+1] == '|' )
					{
						tok = TOKEN.TOK_OR;
						index +=2;
					}
					else 
					{
						tok = TOKEN.ILLEGAL_TOKEN;
						index++;
					}
					break;
				case '/':

					if ( IExpr[index+1] == '/' )
					{
						SkipToEoln();
						goto re_start;
					}
					else 
					{
						tok = TOKEN.TOK_DIV; 
						index++;
					}
					break;
				case '#':
					if ( IExpr[index+1] != '(' ) 
					{
						tok = TOKEN.ILLEGAL_TOKEN; 
						index++;
					}
					else 
					{
						index+=2;

						if (char.IsLetter(IExpr[index]))
						{
							cell_tbl_clm.SheetName  = ExtractString(); 

						}

						if ( IExpr[index] != '.' )
						{
							tok = TOKEN.ILLEGAL_TOKEN; 
							break; 
						}
						index++;
						 
						if (char.IsLetter(IExpr[index]))
						{
							cell_tbl_clm.CellName  = ExtractString(); 

						}

						if ( IExpr[index] != '(' )
						{
							tok = TOKEN.ILLEGAL_TOKEN; 
							break; 
						}
 
						tok = TOKEN.TOK_CELLDBCOLUMN;
						
                       

					}
					break;   
				case '"':
					String x="";
					index++;
					while ( index < length && IExpr[index] !='"' )
					{
						x = x + IExpr[index];
						index ++;
					}

					if ( index == length )
						tok = TOKEN.ILLEGAL_TOKEN; 
					else 
					{
						index++;
						last_str = x;
						tok = TOKEN.TOK_STRING;   
					}
					break;
				case '@':
					if ( IExpr[index+1] != '(' ) 
					{
						tok = TOKEN.ILLEGAL_TOKEN; 
						index++;
					}
					else 
					{
						index+=2;

						if (char.IsLetter(IExpr[index]))
						{
							tbl_clm.TableName  = ExtractString(); 

						}

						if ( IExpr[index] != '.' )
						{
							tok = TOKEN.ILLEGAL_TOKEN; 
							break; 
						}
						index++;
						 
						if (char.IsLetter(IExpr[index]))
						{
							tbl_clm.ColumnName  = ExtractString(); 

						}

						if ( IExpr[index] != ')' )
						{
							tok = TOKEN.ILLEGAL_TOKEN; 
							break; 
						}
 
						tok = TOKEN.TOK_DBCOLUMN;
						index++;
                       

					}
					break;
				case '$':
                   
					if ( IExpr[index+1] != '(' ) 
					{
						tok = TOKEN.ILLEGAL_TOKEN; 
						index++;
					}
					else 
					{
						index+=2;

						if (char.IsLetter(IExpr[index]))
						{
							cell_value.SheetName = ExtractString(); 

						}

						if ( IExpr[index] != '.' )
						{
							tok = TOKEN.ILLEGAL_TOKEN; 
							break; 
						}
						index++;
						 
						if (char.IsLetter(IExpr[index]))
						{
							cell_value.CellName = ExtractString(); 

						}

						if ( IExpr[index] != ')' )
						{
							tok = TOKEN.ILLEGAL_TOKEN; 
							break; 
						}
 
						tok = TOKEN.TOK_CELL;
						index++;
                       


					}
					break;

				case '%':

					if ( IExpr[index+1] != '(' ) 
					{
						tok = TOKEN.ILLEGAL_TOKEN; 
						index++;
					}
					else 
					{
						index+=2;

						if (char.IsLetter(IExpr[index]))
						{
							cell_value.SheetName = ExtractString(); 

						}

						if ( IExpr[index] != '.' )
						{
							tok = TOKEN.ILLEGAL_TOKEN; 
							break; 
						}
						index++;
						 
						if (char.IsLetter(IExpr[index]))
						{
							cell_value.CellName = ExtractString(); 

						}

						if ( IExpr[index] != ')' )
						{
							tok = TOKEN.ILLEGAL_TOKEN; 
							break; 
						}
 
						tok = TOKEN.TOK_CELLNAME;
						index++;
                       


					}
					break;

 
				
				default:
					if ( char.IsDigit(IExpr[index]) )
					{

						String str="";
				  
						while ( index < length &&( IExpr[index] == '0' || 
							IExpr[index]==  '1' ||
							IExpr[index] == '2' || 
							IExpr[index]== '3'  ||
							IExpr[index] == '4' || 
							IExpr[index]== '5'  ||
							IExpr[index] == '6' || 
							IExpr[index]== '7'  ||
							IExpr[index] == '8' || 
							IExpr[index]== '9'  ))
						{
							str += Convert.ToString(IExpr[index]);
							index ++;
						}

						if ( IExpr[index] == '.' )
						{
							str = str +".";
							index++;
							while ( index < length &&( IExpr[index] == '0' || 
								IExpr[index]==  '1' ||
								IExpr[index] == '2' || 
								IExpr[index]== '3'  ||
								IExpr[index] == '4' || 
								IExpr[index]== '5'  ||
								IExpr[index] == '6' || 
								IExpr[index]== '7'  ||
								IExpr[index] == '8' || 
								IExpr[index]== '9'  ))
							{
								str += Convert.ToString(IExpr[index]);
								index ++;
							}

						}




				   
						number = Convert.ToDouble(str); 
						tok = TOKEN.TOK_DOUBLE; 
                

					}
					else if ( char.IsLetter(IExpr[index]))
					{

						String tem = Convert.ToString(IExpr[index]);
						index++;
						while ( index < length && ( char.IsLetterOrDigit(IExpr[index]) ||
							IExpr[index]=='_'))
						{
							tem += IExpr[index];
							index++;
						}

						tem = tem.ToUpper(); 

						for( int i=0; i < keyword.Length ; ++ i )
						{
							if ( keyword[i].Value.CompareTo(tem) == 0 )
								return keyword[i].tok; 

						}


						this.last_str = tem; 
						


						return TOKEN.TOK_UNQUOTED_STRING; 



					}
					else 
					{
						return TOKEN.ILLEGAL_TOKEN; 
					}
					break;
			}

			return tok; 

		}

		/// <summary>
		///  Return the last grabbed number from the steam
		/// </summary>
		/// <returns></returns>
		public 	double GetNumber() 
		{ 
			return number; 
		
		}

	}
	/// <summary>
	/// Summary description for CRuleEngine.
	/// </summary>
	public class CRuleEngine
	{
		private String m_filename;
		private String m_rootpath; 


		/// <summary>
		///    Ctor
		/// </summary>
		/// <param name="Filename"></param>
		/// <param name="rootpath"></param>
		public CRuleEngine(String Filename,String rootpath)
		{
			m_filename = Filename;
			m_rootpath = rootpath;

		}
		/// <summary>
		///   
		/// </summary>
		/// <returns></returns>
		public bool ProcessRulesAndUpdateDB()
		{
			CForeCastUpdate fupd = new CForeCastUpdate(m_filename);
			return fupd.Update();
		}



	}

}
