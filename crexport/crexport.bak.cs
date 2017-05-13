#region Copyright(c) 2003, Teng-Yong Ng
/*
 Copyright(c) 2003, Teng-Yong Ng
  
 This software is provided 'as-is', without any express or implied warranty. In no 
 event will the authors be held liable for any damages arising from the use of this 
 software.
 
  Permission is granted to anyone to use this software for any purpose, including 
 commercial applications, and to alter it and redistribute it freely, subject to the 
 following restrictions:

 1. The origin of this software must not be misrepresented; you must not claim that 
 you wrote the original software. If you use this software in a product, an 
 acknowledgment (see the following) in the product documentation is required.

 Portions Copyright 2003 Teng-Yong Ng

 2. Altered source versions must be plainly marked as such, and must not be 
 misrepresented as being the original software.

 3. This notice may not be removed or altered from any source distribution.
*/

#endregion
/*
file		:crexport.cs
version	:2.0.*
compile	:csc /t:exe /out:crexport.exe /debug+ 
			/r:"H:\Program Files\Common Files\
			Crystal Decisions\1.0\Managed\CrystalDecisions.CrystalReports.Engine.dll" 
			/r:"H:\Program Files\Common Files\
			Crystal Decisions\1.0\Managed\CrystalDecisions.Shared.dll" crexport.cs

Created By: Teng-Yong Ng

Version history
1.1.1152.38285 - First release
1.1.1190.40252 - Bug fixed. Crystal Reports without parameter does not refresh while exporting
1.2.1217.18398 - Allow user to pass multiple values into multiple value parameters 
1.2.1247.16937 - Allow user to pass null value discrete parameter
1.2.1271.40511 - Allow user to specify export html file into single or separated page.
1.2.1419.21168 - zlib/libpng license agreement added.
*/

using System;
using System.Reflection;
using System.Text;
using System.Collections;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;


namespace crexport
{
	public class RptInfo
	{
		#region RptInfo private variables
		private string Username;//Report database login username
		private string Password;//Report database login password
		private string RptPath;//Crystal Report path and filename
		private string OutputPath;//Output file path and filename
		private string ServerName;//Server name of specified crystal Report
		private string DatabaseName;//Database name of specified crystal Report
		private string OutputFormat;//Crystal Report exported format
		private string ExportOption;//Export Option. File,MsMail or Exchange Folder

		private string[] RptParamDisName;//Discrete type parameter name
		private string[] RptParamDisValue;//Discrete type parameter value
		
		private string[] RptParamRangeName;//Range type parameter name
		private string[][] RptParamStartValue;//First value of range type parameter
		private string[][] RptParamEndValue;//Last value of range type parameter
		
		private int ParamDisNumber;//Number of discrete parameters
		private int ParamRangeNumber;//Number of range parameters
		private int[] RangeValueNumber;//Number of multiple range values

		private string SeparatedHTML;//Determine user need export into single or separated html file
		//Microsoft Mail related parameters
		private string ToList;
		private string CcList;
		private string Subject;
		private string Message;
		private string MailUsername;
		private string MailPassword;

		private bool AnyError;
		#endregion private variables

		public RptInfo(string[] parameters)
		{
			int i=0;
			int j=0;
			int k=0;
			int l=0;
			int m=0;
			bool paramDisValid=false;
			bool paramRangeValid=false;
            string defaultFileFormat = "txt";

			this.AnyError=false;
			
			#region RptInfo arguments assignment
			//Assign user parameter to respective members
			foreach (string input in parameters)
			{
				string option=input.Substring(0,2);
				string paraInfo=input.Substring(2,input.Length-2);

				switch (option)
				{
					case "-U":
						this.Username=paraInfo;
						break;
					case "-P":
						this.Password=paraInfo;
						break;
					case "-F":
						this.RptPath=paraInfo;
						break;
					case "-O":
						this.OutputPath=paraInfo;
						break;
					case "-S":
						this.ServerName=paraInfo;
						break;
					case "-D":
						this.DatabaseName=paraInfo;
						break;
					case "-E":
						this.OutputFormat=paraInfo;
						break;
					case "-X":
						this.ExportOption=paraInfo;
						break;
					case "-t":
						this.ToList=paraInfo;
						break;
					case "-c":
						this.CcList=paraInfo;
						break;
					case "-s":
						this.Subject=paraInfo;
						break;
					case "-m":
						this.Message=paraInfo;
						break;
					case "-u":
						this.MailUsername=paraInfo;
						break;
					case "-p":
						this.MailPassword=paraInfo;
						break;
					case "-N"://number of discrete parameters
						try
						{
							ParamDisNumber=Convert.ToInt32(paraInfo);	
							RptParamDisName=new string[ParamDisNumber];
							RptParamDisValue=new string[ParamDisNumber];
							paramDisValid=true;
						}
						catch (FormatException)
						{
							Console.WriteLine("Parameter {0} must be an integer",option);
							AnyError=true;
						}
						catch (Exception)
						{
							Console.WriteLine("Exception caught at parameter N ");
							AnyError=true;
						}
						break;
					case "-M"://number of range parameters
						try
						{
							ParamRangeNumber=Convert.ToInt32(paraInfo);
							RptParamRangeName=new string[ParamRangeNumber];
							RptParamStartValue=new string[ParamRangeNumber][];
							RptParamEndValue=new string[ParamRangeNumber][];
							RangeValueNumber=new Int32[ParamRangeNumber];
							paramRangeValid=true;
						}
						catch (FormatException)
						{
							Console.WriteLine("Parameter {0} must be an integer",option);
							AnyError=true;
						}
						catch (Exception)
						{
							Console.WriteLine("Exception caught at parameter M");
							AnyError=true;
						}

						break;
					case "-A":
					{
						if (paramDisValid)
						{
							try
							{
								this.RptParamDisName[i]=paraInfo;
								i++;
							}
							catch (Exception)
							{
								Console.WriteLine("Exception caught at parameter A");
								AnyError=true;
							}
							
						}
						break;
					}

					case "-B":
					{
						if (paramRangeValid)
						{
							try
							{
								this.RptParamRangeName[m]=paraInfo;
								m++;
							}
							catch (Exception)
							{
								Console.WriteLine("Exception caught at paramter B");
								AnyError=true;
							}
						}
						break;
					}
					case "-J":
					{
						if (paramDisValid)
						{
							try
							{
								this.RptParamDisValue[j]=paraInfo;
								j++;
							}
							catch (Exception)
							{
								Console.WriteLine("Exception caught at parameter J");
								AnyError=true;
							}
						}
						break;
					}
					case "-V"://get values of range parameters
					{
						if (paramRangeValid)
						{
							//get multiple value range parameter	
							//get number of different parameters
							if (paraInfo.IndexOf("|")==-1)
							#region single value range parameter
							{
								RangeValueNumber[k]=1;
								try
								{
									int delimiter=paraInfo.IndexOf(",");
									int leftbracket=paraInfo.IndexOf("(");
									int	rightbracket=paraInfo.IndexOf(")");
									
									RptParamStartValue[k]=new string[] {paraInfo.Substring(leftbracket+1,delimiter-1)};
									RptParamEndValue[l]=new string[] {paraInfo.Substring(delimiter+1,rightbracket-delimiter-1)};
								}
								catch (Exception)
								{
									Console.WriteLine("Exception caught ");
									AnyError=true;
								}
							}
							#endregion single value range parameter
							else //range parameter consists of multiple values
							#region multi value range parameter
							{
								int startSeekPos=0;//start position to seek for "|"	
								ArrayList delimiterPosition=new ArrayList();
								
								while (paraInfo.IndexOf("|",startSeekPos)!=-1)
								{
									delimiterPosition.Add(paraInfo.IndexOf("|",startSeekPos));
									startSeekPos=paraInfo.IndexOf("|",startSeekPos)+1;
								}
								RptParamStartValue[k]=new string[delimiterPosition.Count];
								RptParamEndValue[l]=new string[delimiterPosition.Count];
								
								string tempString;//temporary string designated to store each range value
								int startPos=0;//start position to read value, value will be from startPos to delimiterPosition
								RangeValueNumber[k]=delimiterPosition.Count;
								for (int deliCount=0;deliCount<delimiterPosition.Count;deliCount++)
								{
									tempString=paraInfo.Substring(startPos,(int)delimiterPosition[deliCount]-startPos);
									try
									{
										int delimiter=tempString.IndexOf(",");
										int leftbracket=tempString.IndexOf("(");
										int	rightbracket=tempString.IndexOf(")");
									
										RptParamStartValue[k][deliCount]= tempString.Substring(leftbracket+1,delimiter-1);
										RptParamEndValue[l][deliCount]= tempString.Substring(delimiter+1,rightbracket-delimiter-1);
										startPos=(int)delimiterPosition[deliCount]+1;		
									}
									catch (Exception)
									{
										Console.WriteLine("Exception caught at parameter V");
										AnyError=true;
									}
									
								}
							}
							#endregion multi value range parameter
							k++;
							l++;
						}
						break;
					}
					case "-H"://get value of Enable separated page
						this.SeparatedHTML=paraInfo;
						break;
					case "-?":
						ShowMsg HelpMsg=new ShowMsg(2);
						AnyError=true;//Set to AnyError so to skip load cyrstal report
						break;
					default:
						Console.WriteLine("Invalid parameter \"{0}\" found",option);
						AnyError=true;
						break;
				}
				#endregion RptInfo arguments assignment
			}
            //assigning default parameter values start

            this.ExportOption = this.ExportOption==null? "File":this.ExportOption;


            if (this.OutputPath == null) //if output path and filename is not specified
            {
                if (this.OutputFormat == null)
                    this.OutputFormat = defaultFileFormat;

                string fileExt = "";

                if (this.OutputFormat == "xlsdata")
                    fileExt = "xls";
                else if (this.OutputFormat == "tab")
                    fileExt = "txt";
                else if (this.OutputFormat == "ertf")
                    fileExt = "rtf";
                else
                    fileExt = this.OutputFormat;

                this.OutputPath = String.Format("{0}-{1}.{2}", this.RptPath.Substring(0, this.RptPath.LastIndexOf(".rpt")), DateTime.Now.ToString("yyyyMMddHHmmss"),fileExt);
            }
            else
            {
                //if output path and filename is specified, use file extension to determine output format.
                if (this.OutputFormat == null)
                {
                    int lastIndexDot = this.OutputPath.LastIndexOf(".");
                    string fileExt = this.OutputPath.Substring(lastIndexDot + 1, 3); 
                    
                    //ensure filename extension has 3 char after the dot (.)
                    if ((this.OutputPath.Length == lastIndexDot + 4) && (fileExt=="rtf"||fileExt=="txt"||fileExt=="csv"||fileExt=="pdf"||fileExt=="rpt"||fileExt=="doc"||fileExt=="xls"||fileExt=="xml"||fileExt=="htm"))
                        this.OutputFormat = this.OutputPath.Substring(lastIndexDot + 1, 3);
                    else
                        this.OutputFormat = defaultFileFormat;
                }
            }


            //assigning default parameter values end
		}

		#region RptInfo properties
		
        public string username
		{
			get
			{
				return Username;
			}

		}
		public string password
		{
			get
			{
				return Password;
			}
		}

		public string rptpath
		{
			get
			{
				return RptPath;
			}
		}
		public string outputpath
		{
			get
			{
				return OutputPath;
			}
		}

		public string servername
		{
			get
			{
				return ServerName;
			}
		}

		public string databasename
		{
			get
			{
				return DatabaseName;
			}
		}

		public string outputformat
		{
			get
			{
				return OutputFormat;
			}
		}

		public string exportoption
		{
			get
			{
				return ExportOption;
			}
		}
		public string rptparamdisname(int index)
		{
			return RptParamDisName[index];
		}
		
		public string rptparamdisvalue(int index)
		{
			return RptParamDisValue[index];
		}

		public string rptparamrangename(int index)
		{
			return RptParamRangeName[index];
		}
		
		public string rptparamstartvalue(int index,int subIndex)
		{
			return RptParamStartValue[index][subIndex];
		}

		public void setrptparamstartvalue(string Value,int index,int subIndex)
		{
			RptParamStartValue[index][subIndex]=Value;
		}

		public string rptparamendvalue(int index,int subIndex)
		{
			return RptParamEndValue[index][subIndex];
		}

		public void setrptparamendvalue(string Value,int index,int subIndex)
		{
			RptParamEndValue[index][subIndex]=Value;
		}

		public int paramdisnumber
		{
			get
			{
				return ParamDisNumber;
			}
		}
		public int paramrangenumber
		{
			get
			{
				return ParamRangeNumber;
			}
		}
		public string tolist
		{
			get 
			{
				return ToList;
			}
		}

		public string cclist
		{
			get
			{
				return CcList;
			}
		}
		public string subject
		{
			get
			{
				return Subject;
			}
		}

		public string message
		{
			get
			{
				return Message;
			}
		}
		public string mailusername
		{
			get
			{
				return MailUsername;
			}
		}
		public string mailpassword
		{
			get
			{
				return MailPassword;
			}
		}
		public bool anyerror
		{
			get
			{
				return AnyError;
			}
		}
		public int rangevaluenumber(int index)
		{
			return this.RangeValueNumber[index];
		}

		public string separatedHtml
		{
			get
			{
				return SeparatedHTML;
			}
		}

		#endregion RptInfo properties
	}

	#region Exceptions Classes
	public class NullArgumentException : Exception
	{
		// Base Exception class constructors.
		public NullArgumentException()
			:base() {}
		public NullArgumentException(String message)
			:base(message) {}
		public NullArgumentException(String message, Exception innerException)
			:base(message, innerException) {}
	}

	public class InvalidOutputException: Exception
	{
		public InvalidOutputException()
			:base() {}
		public InvalidOutputException(String message)
			:base(message) {}
		public InvalidOutputException(String message, Exception innerException)
			:base(message, innerException) {}	
	}

	public class InvalidServerException: Exception
	{
		public InvalidServerException()
			:base(){}
		public InvalidServerException(string message)
			:base(message) {}
		public InvalidServerException(string message,Exception innerException)
			:base(message,innerException){}
	}

	public class NullParamNameException: Exception
	{
		public NullParamNameException()
			:base(){}
		public NullParamNameException(string message)
			:base(message) {}
		public NullParamNameException(string message,Exception innerException)
			:base(message,innerException){}
	}
	
	public class NullParamValueException: Exception
	{
		public NullParamValueException()
			:base(){}
		public NullParamValueException(string message)
			:base(message) {}
		public NullParamValueException(string message,Exception innerException)
			:base(message,innerException){}
	}

	public class NullExportTypeException: Exception
	{
		public NullExportTypeException()
			:base(){}
		public NullExportTypeException(string message)
			:base(message) {}
		public NullExportTypeException(string message,Exception innerException)
			:base(message,innerException){}
	}
	#endregion Exceptions Classes

	public class ShowMsg
	{
		public ShowMsg(byte mode)
		{
            if (mode==0)
			{
				AssemblyName aName=GetType().Assembly.GetName();
				Console.WriteLine("{0} (version: {1})",aName.Name,aName.Version);
				Console.WriteLine("Copyright(c) 2011 Rainforest Software Solution (http://www.rainforestnet.com)\n");
			}
			else if (mode==1)
			{
				Console.WriteLine("Type \"crexport -?\" for help");
			}
			else if (mode==2)
			{
				AssemblyName aName = GetType().Assembly.GetName();
				Console.WriteLine("{0} (version: {1})",aName.Name,aName.Version);
				Console.WriteLine("Copyright(c) 2011 Rainforest Software Solution  http://www.rainforestnet.com\n");
				Console.WriteLine("Crystal Report Exporter Paramters Listing");
				Console.WriteLine("---------------------------------------------------");
				Console.WriteLine("-U database login username");
				Console.WriteLine("-P database login password");
				Console.WriteLine("-F Crystal report filename");
				Console.WriteLine("-O Crystal report Output filename");
				Console.WriteLine("-E Export file type.(pdf,doc,xls,rtf,htm,rpt)");
				Console.WriteLine("-S Server Name");
				Console.WriteLine("-D Database Name");
				Console.WriteLine("-X Export Type (File,Msmail)");
				Console.WriteLine("-N Number of Discreate Parameters");
				Console.WriteLine("-A Discreate parameter name");
				Console.WriteLine("-J Discreate parameter value");
				Console.WriteLine("-M Number of range parameters");
				Console.WriteLine("-B range parameter name");
				Console.WriteLine("-V range parameter value([start value],[end value])");
				Console.WriteLine("-H Enable separated HTML page option(y or n).Only applicable if export to html file");
				//Microsoft Mail related parameter
				Console.WriteLine("Microsoft Mail related parameters");
				Console.WriteLine("-t Mail \"To\" List");
				Console.WriteLine("-c Mail \"CC\" List");
				Console.WriteLine("-s Mail Subject");
				Console.WriteLine("-m Mail Message");
				Console.WriteLine("-u Mail Account username");
				Console.WriteLine("-p Mail Account password");
				Console.WriteLine("---------------------------------------------------");
				Console.WriteLine("Example: crexport -Uuser1 -Pmypass -S\"Extreme Sample Database\" -Fc:\\test.rpt -Od:\\test.rtf -Epdf -XFile -N1 -Asupplier -J\"Active Outdoors\" -M1 -Bdaterange -V(12-01-2001,12-04-2002)");
				Console.WriteLine("\nNote: -N and -M must be defined before defining -A,-J and -B,-V parameters");
			}
		}
	}

	class crexport
		{
			/// <summary>
			/// The main entry point for the application.
			/// </summary>
			[STAThread]
			static void Main(string[] args)
			{
				try
				{

					RptInfo ReportInfo =new RptInfo(args);
					if (!ReportInfo.anyerror)
					{
						ReportDocument Report = new ReportDocument();
						ShowMsg WelcomeMsg=new ShowMsg(0);
						try
						{
							
							if (args.Length==0)
							{
								throw new
									NullArgumentException("No parameter found!");
							}
							if (ReportInfo.outputpath==null)
							{
								throw new
									InvalidOutputException("Unspecified Output filename");
							}
							if (ReportInfo.servername==null)
							{
								throw new
									InvalidServerException("Unspecified Server Name");
							}
							if (ReportInfo.exportoption==null)
							{
								throw new
									NullExportTypeException("Unspecified Export Type");
							}
						
							Report.Load(ReportInfo.rptpath,OpenReportMethod.OpenReportByTempCopy);
							TableLogOnInfo logonInfo= new TableLogOnInfo();
							foreach (Table table in Report.Database.Tables)
							{
								logonInfo.ConnectionInfo.ServerName = ReportInfo.servername;
								logonInfo.ConnectionInfo.DatabaseName = ReportInfo.databasename;
								if (ReportInfo.username!=null)
								{
									logonInfo.ConnectionInfo.UserID = ReportInfo.username;
								}
								if (ReportInfo.password!=null)
								{
									logonInfo.ConnectionInfo.Password = ReportInfo.password;
								}
								table.ApplyLogOnInfo(logonInfo);
							}
							//set the export format
                            if (ReportInfo.outputformat.ToUpper() == "RTF")
                                Report.ExportOptions.ExportFormatType = ExportFormatType.RichText;
                            else if (ReportInfo.outputformat.ToUpper() == "TXT")
                                Report.ExportOptions.ExportFormatType = ExportFormatType.Text;
                            else if (ReportInfo.outputformat.ToUpper() == "TAB")
                                Report.ExportOptions.ExportFormatType = ExportFormatType.TabSeperatedText;
                            else if (ReportInfo.outputformat.ToUpper() == "CSV")
                                Report.ExportOptions.ExportFormatType = ExportFormatType.CharacterSeparatedValues;
                            else if (ReportInfo.outputformat.ToUpper() == "PDF")
                                Report.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                            else if (ReportInfo.outputformat.ToUpper() == "RPT")
                                Report.ExportOptions.ExportFormatType = ExportFormatType.CrystalReport;
                            else if (ReportInfo.outputformat.ToUpper() == "DOC")
                                Report.ExportOptions.ExportFormatType = ExportFormatType.WordForWindows;
                            else if (ReportInfo.outputformat.ToUpper() == "XLS")
                                Report.ExportOptions.ExportFormatType = ExportFormatType.Excel;
                            else if (ReportInfo.outputformat.ToUpper() == "XLSDATA")
                                Report.ExportOptions.ExportFormatType = ExportFormatType.ExcelRecord;
                            else if (ReportInfo.outputformat.ToUpper() == "ERTF")
                                Report.ExportOptions.ExportFormatType = ExportFormatType.EditableRTF;
                            else if (ReportInfo.outputformat.ToUpper() == "XML")
                                Report.ExportOptions.ExportFormatType = ExportFormatType.Xml;
                            else if (ReportInfo.outputformat.ToUpper() == "HTM")
                            {
                                int i = 1;
                                while (ReportInfo.outputpath.IndexOf(@"\", (ReportInfo.outputpath.Length) - i, i) < 0)
                                {
                                    i++;
                                }
                                int lastSlashPos = ReportInfo.outputpath.Length - i + 1; //Last back slash position in outputpath string
                                string baseFolder = ReportInfo.outputpath.Substring(0, lastSlashPos - 1);

                                HTMLFormatOptions htmlFormatOptions = new HTMLFormatOptions();
                                htmlFormatOptions.HTMLBaseFolderName = baseFolder;
                                htmlFormatOptions.HTMLFileName = ReportInfo.outputpath;
                                if (ReportInfo.separatedHtml == "n")
                                {
                                    htmlFormatOptions.HTMLEnableSeparatedPages = false;
                                }
                                else
                                {
                                    htmlFormatOptions.HTMLEnableSeparatedPages = true;
                                }
                                htmlFormatOptions.HTMLHasPageNavigator = true;
                                htmlFormatOptions.FirstPageNumber = 1;

                                Report.ExportOptions.ExportFormatType = ExportFormatType.HTML40;
                                Report.ExportOptions.FormatOptions = htmlFormatOptions;
                            }
		
			
							//Parameters
							ParameterDiscreteValue paraValue;
							ParameterRangeValue paraRangeValue;
							ParameterFieldDefinition ParamDef;
							ParameterFieldDefinitions ParamDefs;
							ParameterValues ParamValues;

							ParamDefs= Report.DataDefinition.ParameterFields;
							if ((ReportInfo.paramdisnumber==0) & (ReportInfo.paramrangenumber==0))
							{
							   Report.Refresh();
							}
							#region Pass discrete parameter to ParameterDefiniton object
							for (int i=0;i<ReportInfo.paramdisnumber;i++)
							{
								
								if (ReportInfo.rptparamdisname(i)==null)
								{
									throw new NullParamNameException("Null Discrete Parameter found");
								}
								ParamDef=ParamDefs[ReportInfo.rptparamdisname(i)];
								ParamValues=new ParameterValues();
						
								//Pass single or multiple value into one Discrete parameter
								//paramList consists of single of multiple values of discrete parameter
								string paramList=ReportInfo.rptparamdisvalue(i);
								
								if (paramList.IndexOf("|")==-1)
								{
									//if | character is not found is parameter values
									Console.WriteLine("Discrete Parameter \"{0}\" = \"{1}\"",ReportInfo.rptparamdisname(i),paramList);
									paraValue=new ParameterDiscreteValue();
									//This part is added to support NULL value parameter
									if (paramList=="")
									{
										paraValue.Value=null;
									}
									else
									{
										paraValue.Value=paramList;
									}
									ParamValues.Add(paraValue);
									ParamDef.ApplyCurrentValues(ParamValues);
								}
								else
								{
									//if "|" found,means multiple values parameter found
									int disStart=0;
									string singleDisValue;
			
									ParamDef.EnableAllowMultipleValue=true;
									//disStart is start search position of parameter string
									while (paramList.IndexOf("|",disStart)!=-1)
									{
										singleDisValue=paramList.Substring(disStart,paramList.IndexOf("|",disStart+1)-disStart);
										Console.WriteLine("Discrete Parameter \"{0}\" = \"{1}\"",ReportInfo.rptparamdisname(i),singleDisValue);
										paraValue=new ParameterDiscreteValue();
										if (singleDisValue=="")
										{
											paraValue.Value=null;
										}
										else
										{
											paraValue.Value=singleDisValue;
										}
										ParamValues.Add(paraValue);
										disStart=paramList.IndexOf("|",disStart)+1;
									}
									ParamDef.ApplyCurrentValues(ParamValues);
								}
							}
							#endregion Pass discrete parameter to ParameterDefinition object
							#region Pass range value to ParameterDefinition object
							for (int i=0;i<ReportInfo.paramrangenumber;i++)
							{
								if (ReportInfo.rptparamrangename(i)==null)
								{
									throw new NullParamNameException("Null Range Parameter found");
								}
								ParamDef=ParamDefs[ReportInfo.rptparamrangename(i)];
								ParamValues=new ParameterValues();
								//Single value range parameter
								for (int v=0;v<ReportInfo.rangevaluenumber(i);v++)
								{
									paraRangeValue=new ParameterRangeValue();
									if (ReportInfo.rptparamstartvalue(i,v)=="")
									{
										paraRangeValue.StartValue=null;
									}
									else
									{
										paraRangeValue.StartValue=ReportInfo.rptparamstartvalue(i,v);
									}	
									
									if (ReportInfo.rptparamendvalue(i,v)=="")
									{
										paraRangeValue.EndValue=null;
									}
									else
									{
										paraRangeValue.EndValue=ReportInfo.rptparamendvalue(i,v);
									}
									Console.WriteLine("Range Parameter \"{0}\" from {1} to {2}",ReportInfo.rptparamrangename(i),ReportInfo.rptparamstartvalue(i,v),ReportInfo.rptparamendvalue(i,v));
									ParamValues.Add(paraRangeValue);
								}
								ParamDef.ApplyCurrentValues(ParamValues);
							}	
							#endregion Pass range value to ParameterDefinition object

							if (ReportInfo.exportoption.ToUpper()=="FILE")
							{
								Report.ExportOptions.ExportDestinationType =ExportDestinationType.DiskFile;
								//DiskFileDestinationOptions
								DiskFileDestinationOptions diskOpts = new DiskFileDestinationOptions();
								diskOpts.DiskFileName = ReportInfo.outputpath;
								Report.ExportOptions.DestinationOptions = diskOpts;
							}
							else if (ReportInfo.exportoption.ToUpper()=="MSMAIL")
							{
								Report.ExportOptions.ExportDestinationType =ExportDestinationType.MicrosoftMail;
								MicrosoftMailDestinationOptions msmailOpts=new MicrosoftMailDestinationOptions();
								msmailOpts.MailToList=ReportInfo.tolist;
								msmailOpts.MailCCList=ReportInfo.cclist;
								msmailOpts.MailSubject=ReportInfo.subject;
								msmailOpts.MailMessage=ReportInfo.message;
								msmailOpts.UserName=ReportInfo.mailusername;
								msmailOpts.Password=ReportInfo.mailpassword;
								Report.ExportOptions.DestinationOptions=msmailOpts;
							}
							else
							{
								throw new NullExportTypeException("\nError: Invalid Export Type");
							}
							
							Report.Export ();

						}   
						catch (NullArgumentException Er)
						{
							Console.WriteLine("\nError: "+Er.Message);
							ShowMsg HelpMsg=new ShowMsg(1);
						}

						catch (InvalidOutputException Er)
						{
							Console.WriteLine("\nError: "+Er.Message);
							ShowMsg HelpMsg=new ShowMsg(1);
						}
					
						catch (InvalidServerException Er)
						{
							Console.WriteLine("\nError: "+Er.Message);
							ShowMsg HelpMsg=new ShowMsg(1);
						}
					
						catch (LogOnException)
						{
							Console.WriteLine("\nError: Failed to logon to Database. Check username, password, server name and database name parameter");
						}

						catch (LoadSaveReportException)
						{
							Console.WriteLine("\nError: Failed to Load or save crystal report");
						}
					
						catch (NullParamNameException Er)
						{
							Console.WriteLine("\nError: {0}",Er.Message);
						}
					
						catch (NullExportTypeException Er)
						{
							Console.WriteLine("\nError: {0}",Er.Message);
						}

						catch (Exception Er)
						{
							Console.WriteLine("\nMisc Error: {0}",Er.Message);
							ShowMsg HelpMsg=new ShowMsg(1);
						}
						finally
						{
							Report.Close();
						}
					}
				}
				catch (Exception Er)
				{
					Console.WriteLine("\n System Error: {0}",Er.Message);
				}

			}
		}
	}
