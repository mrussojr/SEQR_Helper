/*
 * Created by SharpDevelop.
 * User: russom
 * Date: 7/12/2016
 * Time: 8:59 AM
 * 
 * SEQR Helper
 * Used to move project files to Albany SEQR Clearinghouse drive, compare
 * files on Local SEQR drive and Albany SEQR Clearinghouse drive, and create
 * project files on the Local SEQR drive.
 */
using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace SEQR_Helper
{
	class Program
	{
		#region Menu Classes
		/*
		 * The MenuChoice class was created to be used for creating and using the Park and Project menus.
		 * I needed a way to store both an id for selection and a name and later be able to sort them.
		 */
		public class MenuChoice {
			public int id;
			public string name;
			
			public MenuChoice (int i, string n)
			{
				id = i;
				name = n;
			}
		}
		
		/*
		 * The MenuChoiceList was created to store menu items, used for sorting menu items, and retrieving
		 * the name and/or id of a menu item
		 */ 
		public class MenuChoiceList : List<MenuChoice> {
			public MenuChoiceList(){}
			//checks if a MenuChoice with the given id exists
			public bool IdInList(int id) {
				foreach(MenuChoice m in this)
				{
					if(m.id == id)
						return true;
				}
				
				return false;
			}
			//checks if a MenuChoice with the given name exists
			public bool NameInList(string name) {
				foreach(MenuChoice m in this)
				{
					if(m.name == name)
						return true;
				}
				
				return false;
			}
			//returns the name of the MenuChoice with the given id
			public string NameById(int id) {
				foreach (MenuChoice m in this) {
					if(m.id == id)
						return m.name;
				}
				
				return "";
			}
		}
		#endregion
		
		#region Properties
		private static MenuChoiceList parksMenu = new MenuChoiceList(); //list to hold the parks menu
		private static MenuChoiceList projectsMenu = new MenuChoiceList(); //list to hold the projects menu
		private static MenuChoiceList regionMenu = new MenuChoiceList(); //list to hold the regions menu
		//location of the configuration file which holds the paths for working with projects
		private static string configFilePath = System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Replace(@"file:///", "").Replace("SEQR_Helper.exe", "") + "config.cfg";		
		//strings used in config file and in prompts
		private static string lPathName = "Local SEQR drive";
		private static string tPathName = "Project Template file";
		private static string aPathName = "Albany SEQR drive";
		//strings to hold paths once loaded
		private static string localPath;
		private static string templatePath;		
		private static string albPath;
		#endregion
		
		public static void Main(string[] args)
		{	
			ChangeSettings();
			LoadPaths();
			GenList("Parks", 0);
			Console.WriteLine("Welcome to the SEQR Helper program!");
			
			int menuChoice = 0;
			
			//following used in the event command line parameters are used with program
			if (args.Length > 0)
				Int32.TryParse(args[0], out menuChoice);
			
			MenuSwitcher(menuChoice);
		}
		
		#region Console and program settings
		//adjusts the console settings		
		private static void ChangeSettings()
		{
			Console.WindowHeight = 35;
			Console.WindowWidth = 160;
			Console.Title = "SEQR Helper - " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
			Console.BackgroundColor = ConsoleColor.DarkCyan;
			Console.ForegroundColor = ConsoleColor.Black;
		}
		
		//loads paths from config file
		private static void LoadPaths() {
			string[] paths;
	
			//prompt user for paths if config file doesn't exist or is empty
			if(!File.Exists(configFilePath) || File.ReadAllText(configFilePath).Equals("")) {
				Console.WriteLine("Welcome to SEQR Helper!\nSince this is your first time running the application we need to set up a few things.");
				WritePaths(true, true, true);
			}

			paths = ReadPaths();

			localPath = paths[0];
			templatePath = paths[1];
			albPath = paths[2];	
		}
		#endregion
		
		#region Menus
		//prints the main menu
		private static void MainMenu()
		{
			int menuChoice = 0;
			
			Console.WriteLine("Please choose from one of the following options:");
			Console.WriteLine("1. Copy files from local drive to albany drive.");
			Console.WriteLine("2. Compare files on local and albany drives.");
			Console.WriteLine("3. Create project file on local drive.");
			Console.WriteLine("4. Check SEQR drive and Template File paths.");
			Console.WriteLine("5. Exit SEQR Helper.");
			
			Int32.TryParse(Console.ReadLine(), out menuChoice);
			
			MenuSwitcher(menuChoice);
		}
		
		//prints the submenu
		private static void SubMenu(string menuType)
		{
			int parkChoice = -1;
			
			//loop until user chooses park in list or 0
			while(!parksMenu.IdInList(parkChoice)) {
				if(parkChoice == 0)
				{
					ReturnToMainMenu(false);
					return;
				}
				Console.WriteLine("Please select a valid park!");
				parkChoice = MenuGen("Park");
			}
						
			GenList("Projects", parkChoice);
			
			if(!menuType.Equals("Create")) {
				if(projectsMenu.Count == 0){
					Console.WriteLine("No projects currently exist, press any key to return to the main menu to create a project");
					ReturnToMainMenu(true);
				}else{
					int projectChoice = -1;
					
					//loop until user chooses project in list or 0
					while(!projectsMenu.IdInList(projectChoice)) {
						if(projectChoice == 0)
						{
							Console.Clear();
							SubMenu(menuType);
							return;
						}
						Console.WriteLine("Please select a valid project!");
						projectChoice = MenuGen("Project");
					}
					
					if(menuType.Equals("Copy"))
					{
						CopyFiles(parkChoice, projectChoice);
					}
					else if(menuType.Equals("Compare"))
					{
						CompareFiles(parkChoice, projectChoice, false);
					}
				}
			}
			else
			{
				CreateProject(parkChoice);
			}	
		}
		
		//chooses where to go from main menu
		private static void MenuSwitcher(int menuChoice) {
			switch (menuChoice) {
				case 1:
					SubMenu("Copy");
					break;
				case 2:
					SubMenu("Compare");
					break;
				case 3:
					SubMenu("Create");
					break;
				case 4:
					PathsMenu();
					break;
				case 5:
					Console.WriteLine("Goodbye...");
					break;
				case 99://TODO: Get rid of this case when done testing
					Console.Clear();
					Console.WriteLine("Secret Mode Enabled");
					EditChecklist("Bowman Lake", "Test Project");
					break;
				default:
					ReturnToMainMenu(false);
					break;
			}
		}
		
		//generates either the park menu or project menu based on arguments, returns users choice
		private static int MenuGen(string type) {
			int choice;
			MenuChoiceList menu;
			
			if(type.Equals("Park"))
				menu = parksMenu;
			else
				menu = projectsMenu;
				
			Console.Clear();
			Console.WriteLine("Please choose the {0} or 0 to return to the previous menu:", type);
			
			foreach(MenuChoice m in menu)
			{
				Console.Write("{0}. {1}\n", m.id, m.name);
			}
			
			try
			{
				choice = Int32.Parse(Console.ReadLine());
			}
			catch(Exception)
			{
				choice = MenuGen(type); //recursively catch menu choice when invalid selection is made
			}
			
			return choice;
		}
		
		//creates the paths menu and performs desired operations
		private static void PathsMenu() {
			bool updated = true;
			string[] paths = ReadPaths();
			int choice = -1;
			
			Console.Clear();
			Console.WriteLine("Current paths are: ");			
			Console.WriteLine("1. {0}", paths[0]);
			Console.WriteLine("2. {0}", paths[1]);
			Console.WriteLine("3. {0}", paths[2]);
			Console.WriteLine("Choose the path to edit or 0 to return to the main menu");
			
			Int32.TryParse(Console.ReadLine(), out choice);
				
			switch (choice) {
				case 0:
					updated = false;
					ReturnToMainMenu(false);
					break;
				case 1:
					WritePaths(true, false, false);
					break;
				case 2:
					WritePaths(false, true, false);
					break;
				case 3:
					WritePaths(false, false, true);
					break;
				default:
					updated = false;
					PathsMenu();  //recursively print path menu on invalid selection
					break;
			}				
			
			if(updated)
			{
				GenList("Parks", 0);
				Console.WriteLine("Path updated successfully, press Enter to continue...");
				ReturnToMainMenu(true);
			}
			
			return;
		}
		#endregion
		
		#region List Generator
		//generates the park and project lists to be used in the menus		
		private static void GenList(string listType, int parkId) {
			string parkName = "";
			string item = "";
			string slash = "";
				
			MenuChoiceList menu = parksMenu;
			
			//sets the information required to generate the project list
			if(listType.Equals("Projects")) {
				parkName = parksMenu.NameById(parkId);
				menu = projectsMenu;
				slash = "\\";
			}
			
			int x = 1;

			if(Directory.Exists(localPath + parkName)) {
				string[] files = Directory.GetDirectories(localPath + parkName);
				
				menu.Clear();
				
				foreach (String f in files)
				{
					item = f.Replace(localPath + parkName + slash, "");
					if(!menu.NameInList(item))
					{
						menu.Add(new MenuChoice(x, item));
						x++;
					}
				}
			}else{
				Console.Clear();
				Console.WriteLine("****There appears to be a problem with your path settings!****\nPlease use the menu option to view / edit your paths.\nPress enter to continue...\n\n\n\n");
				Console.ReadLine();
			}
		}
		
		//generates the list of regions
		private static void GenRegionList() {
			regionMenu.Add(new MenuChoice(1, "Allegany"));
			regionMenu.Add(new MenuChoice(2, "Central"));
			regionMenu.Add(new MenuChoice(3, "Finger Lakes"));
			regionMenu.Add(new MenuChoice(4, "Genesee"));
			regionMenu.Add(new MenuChoice(5, "Long Island"));
			regionMenu.Add(new MenuChoice(6, "New York City"));
			regionMenu.Add(new MenuChoice(7, "Niagara"));
			regionMenu.Add(new MenuChoice(8, "Palisades"));
			regionMenu.Add(new MenuChoice(9, "Saratoga"));
			regionMenu.Add(new MenuChoice(10, "Taconic"));
			regionMenu.Add(new MenuChoice(11, "Thousand Islands"));
		}
		#endregion
		
		#region Major Functions
		//copys a project file from the local drive to the albany drive
		public static void CopyFiles(int parkId, int projectId) {
			string parkName = parksMenu.NameById(parkId);
			string projectName = projectsMenu.NameById(projectId);
			string subPath = parkName + "\\" + projectName;
			
			//does project already exist
			if(Directory.Exists(albPath+subPath)) {
				Console.WriteLine("****Project already exists on Albany drive, use compare option instead.****\nPress enter to continue...");
				ReturnToMainMenu(true);
				return;
			}

			Directory.CreateDirectory(albPath+subPath);
			
			//loop through each folder on local drive and create it on the albany drive			
			foreach(string s in Directory.GetDirectories(localPath+subPath, "*", SearchOption.AllDirectories)) {
				string curDir = s.Replace(localPath+subPath, "");
				Directory.CreateDirectory(albPath+subPath+curDir);
			}
			
			//loop through each file on local drive and copy it to the albany drive
			foreach (string s in Directory.GetFiles(localPath+subPath, "*", SearchOption.AllDirectories)) {
				File.Copy(s, albPath+subPath+s.Replace(localPath+subPath, ""));
			}
			
			AddToWorkbook(parkName, projectName, false);
			
			Console.WriteLine("****Copy to Albany drive complete!****\nPress enter to continue...");
			ReturnToMainMenu(true);
		}
		
		//compares project folders and files on local and albany drives
		public static void CompareFiles(int parkId, int projectId, bool skip) {
			string parkName = parksMenu.NameById(parkId);
			string projectName = projectsMenu.NameById(projectId);
			string subPath = parkName + "\\" + projectName;
			string localFullPath = localPath+subPath;
			string albFullPath = albPath+subPath;
			
			//condition created to recursively compare files and not folders
			if(!skip) {
				//does project exist on albany drive
				if(!Directory.Exists(albPath+subPath)) {
					Console.WriteLine("\n\n\n****Project does not exist on Albany drive, use copy option instead.****\nPress any key to continue...");
					ReturnToMainMenu(true);
					return;
				}
				
				List<string> localDirs = PathRemover(Directory.GetDirectories(localFullPath, "*", SearchOption.AllDirectories), localFullPath);
				List<string> albDirs = PathRemover(Directory.GetDirectories(albFullPath, "*", SearchOption.AllDirectories), albFullPath);
	
				//compare the two directories
				CompareLists(ref localDirs, ref albDirs);
				
				Console.Clear();
				
				if(localDirs.Count == 0 && albDirs.Count == 0)
					Console.WriteLine("Folders on both the Local and Albany drives are the same.\n\n");
				else {
					if(localDirs.Count > 0)
						PrintCompareDirs("Local", localDirs, albFullPath); //gives ability to copy directory to new location
					if(albDirs.Count > 0)
						PrintCompareDirs("Albany", albDirs, localFullPath); //gives ability to copy directory to new location
				}
			}
			
			List<string> localFiles = PathRemover(Directory.GetFiles(localFullPath, "*", SearchOption.AllDirectories), localFullPath);
			List<string> albFiles = PathRemover(Directory.GetFiles(albFullPath, "*", SearchOption.AllDirectories), albFullPath);
			
			//condition used in recursively comparing only files ending when files are the same in both directories
			if(CompareFilesLists(localFiles, localFullPath, albFiles, albFullPath) == 1)
				CompareFiles(parkId, projectId, true);
			else {
				Console.WriteLine("\n\nPress enter to continue....");
				ReturnToMainMenu(true);
			}
		}
		
		//creates a project on the local drive
		private static void CreateProject(int parkId) {
			string parkName = parksMenu.NameById(parkId);
			string fullPath;
			string projectName = "";

			bool flag = false;
			
			//loop until project name is correct or 0
			while(!flag) {
				Console.Clear();
				Console.WriteLine("Please type the name of the project that you'd like to create, type 0 to exit.");
				projectName = Console.ReadLine();
				
				if(projectName.Equals("0")) {
					ReturnToMainMenu(false);
					return;
				}
				
				Console.WriteLine("You entered '{0}' as the project name, is this correct? Y or N", projectName);
				
				if(Console.ReadLine().ToUpper().Equals("Y"))
					flag = true;
			}
			
			fullPath = localPath+parkName+"\\"+projectName;
			
			if(Directory.Exists(fullPath)) {
				Console.WriteLine("****This project already exists!****\nPress any key to try again...");
				Console.ReadLine();
				CreateProject(parkId);
			} else {
				Directory.CreateDirectory(fullPath);
			
				//loop through each folder in the template folder and create it in new project
				foreach(string s in Directory.GetDirectories(templatePath, "*", SearchOption.AllDirectories))
					Directory.CreateDirectory(fullPath + "\\" + s.Replace(templatePath, ""));
			
				//loop through each file in the template folder and copy it to the new project
				foreach(string s in Directory.GetFiles(templatePath, "*", SearchOption.AllDirectories))
					File.Copy(s, fullPath + "\\" + s.Replace(templatePath,""));
				
				Console.WriteLine("****Project created successfully!****\nWould you like to add this project to the SEQR Workbook? Y or N");
				
				EditPPT(parkName, projectName);
				EditChecklist(parkName, projectName);
				
				if(Console.ReadLine().ToUpper().Equals("Y")) {
					AddToWorkbook(parkName, projectName, true);
					Console.WriteLine("****Project added to SEQR Workbook complete!****\nPress enter to return to the main menu.");
					ReturnToMainMenu(true);
				}
				else {
					Console.WriteLine("Press enter to continue...");
					ReturnToMainMenu(true);
				}
			}
		}
		#endregion
		
		#region Helper Functions
		//removes the unnecessary portion of the paths
		public static List<string> PathRemover(string[] list, string pathToRemove)
		{
			List<string> rList = new List<string>();
			
			foreach (string s in list)
				rList.Add(s.Replace(pathToRemove+"\\", ""));
			
			return rList;
		}
		
		//compares 2 lists of directories, removing those that exist in each location
		public static void CompareLists(ref List<string> ar1, ref List<string> ar2) {
			List<string> removals = new List<string>();
			
			foreach (string s in ar1) {
				if(ar2.Contains(s))
					removals.Add(s);
			}
			
			foreach (string s in removals) {
				ar1.Remove(s);
				ar2.Remove(s);
			}
		}
		
		//compares 2 lists of files until all files are the same, returns 1 if there are still 
		//files to compare and 0 when all files are the same
		public static int CompareFilesLists(List<string> ar1, string path1, List<string> ar2, string path2) {
			FileInfo f;
			List<FileInfo> files = new List<FileInfo>();
			
			if(ar1.Count == 0 && ar2.Count == 0)
			{
				Console.WriteLine("****Files are the same on both SEQR drives.****");
				return 0;
			}
			
			ar1.Sort();
			ar2.Sort();
			
			int x = 1;
			
			Console.WriteLine("****Local Files: Files that exist on Local SEQR drive only.****");
			FileLooper(ar1, path1, ar2, path2, ref files, ref x);
			
			Console.WriteLine("\n\n\n****Albany Files: Files that exist on Albany SEQR drive only.****");
			FileLooper(ar2, path2, ar1, path1, ref files, ref x);
			
			if(files.Count > 0) {
				Console.WriteLine("\n\n\n****Choose the file you'd like to move or copy or 0 to return to the main menu.");
				int fileIndex = 999;
				
				Int32.TryParse(Console.ReadLine(), out fileIndex);
				
				if(fileIndex == 0) //returns to main menu without copying files
					return 0;
				else if(fileIndex-1 < files.Count) //confirms valid file selection
				{
					f = files[fileIndex-1];
					
					//checks if file is local or albany before copying to opposite drive
					if(f.FullName.Contains(localPath))
						f.CopyTo(f.FullName.Replace(path1, path2), true);
					else
						f.CopyTo(f.FullName.Replace(path2, path1), true);
				}
				else
				{
					Console.Clear();
					Console.WriteLine("You've made an invalid selection, returning...");
				}
				
				Console.Clear();
				return 1;
			}
			else
				Console.WriteLine("\n\nFiles are the same on both SEQR drives.");
			
			return 0;
		}
		
		//loops through files on 2 drives and compares them
		private static void FileLooper(List<string> ar1, string path1, List<string> ar2, string path2, ref List<FileInfo> files, ref int x) {
			FileInfo f, f2;
			//loops through the files on one drive
			foreach (string s in ar1) {
				f = new FileInfo(path1 + "\\" + s);
				
				//if the file exists on the other drive, it needs to compare the date and size of the file
				if(ar2.Contains(s))
				{
					f2 = new FileInfo(path2 + "\\" + s);
					
					//output is file that exists on both drives, but is of a different size or written to on a different date
					if(f.Length != f2.Length && f.LastWriteTime != f2.LastWriteTime) {
						Console.WriteLine("{0}. {1,-50}{2,-50}{3,-1}bytes", x, s, f.LastWriteTime, f.Length);
						files.Add(f);
						x++;
					}
						
				}
				else //output is file that doesn't exist on other drive
				{
					Console.WriteLine("{0}. {1,-50}{2,-50}{3,-1}bytes", x, s, f.LastWriteTime, f.Length);
					files.Add(f);
					x++;
				}
			}
		}
		
		//loops through folders in directory with option to copy
		public static void PrintCompareDirs(string drive, List<string> dirList, string fullPath) {
			string drive2, prompt;
			
			//sets the drives for looping and copying
			if(drive.Equals("Albany"))
				drive2 = "Local";
			else
				drive2 = "Albany";
			
			Console.WriteLine("****{0} Folders: Folders that exist on the {0} SEQR drive only.****", drive);
			dirList.Sort();
			
			//loop through folders in directory
			foreach(string s in dirList) {
				Console.WriteLine(s);
				Console.WriteLine("Do you want to copy this folder to the {0} SEQR drive? Y or N", drive2);
				prompt = Console.ReadLine();
				
				if(prompt.ToUpper().Equals("Y"))
					Directory.CreateDirectory(fullPath+"\\"+s);
			}
			
			Console.WriteLine("\n\n");
		}
		
		//returns to the main menu, argument determines if console should wait for user input before returning
		public static void ReturnToMainMenu(bool read)
		{
			if(read)
				Console.ReadLine();
			
			Console.Clear();
			MainMenu();
		}
		
		//reads the paths from the config file
		private static string[] ReadPaths() {
			string[] paths = new string[3];
			
			using(StreamReader files = new StreamReader(configFilePath))
			{
				paths[0] = files.ReadLine().Replace(lPathName + "::", "");
				paths[1] = files.ReadLine().Replace(tPathName + "::", "");
				paths[2] = files.ReadLine().Replace(aPathName + "::", "");
			}
			
			return paths;
		}
		
		//writes the paths to the config file, arguments determine which path to write and prompts to display
		private static void WritePaths(bool lPathWrite, bool tPathWrite, bool aPathWrite){
			string lPath, tPath, aPath;
			
			if(lPathWrite)
				lPath = LoadHelper(lPathName, false);
			else
				lPath = localPath;
			
			if(tPathWrite)
				tPath = LoadHelper(tPathName, false);
			else
				tPath = templatePath;
			
			if(aPathWrite)
				aPath = LoadHelper(aPathName, true);
			else
				aPath = albPath;

			string[] fileLines = 
				{
					lPathName + "::" + lPath,
					tPathName + "::" + tPath,
					aPathName + "::" + aPath
				};
		
			File.WriteAllLines(configFilePath, fileLines);
			
			localPath = lPath;
			templatePath = tPath;
			albPath = aPath;
		}
		
		//displays the prompts for the user to input the file paths
		private static string LoadHelper(string prompt, bool selectable) {
			bool flag = false;
			string pathVariable = "";
			
			//for written paths
			if(!selectable) {
				//loop until path typed and confirmed correct
				while(!flag) {
					Console.WriteLine("Please type the path of your {0}:", prompt);
					pathVariable = Console.ReadLine();
					Console.WriteLine("You've typed '{0},' is this correct? Y or N", pathVariable);
					
					if(Console.ReadLine().ToUpper().Equals("Y"))
						flag = true;
				}
			}
			else //for selecting the region since all users have same path for albany drive
			{
				int selection = 0;
				GenRegionList();
				
				//loop until region selected and confirmed correct
				while(!flag) {
					Console.Clear();
					Console.WriteLine("Please select your Region:");
					
					foreach(MenuChoice m in regionMenu)
					{
						Console.Write("{0}. {1}\n", m.id, m.name);
					}
					
					Int32.TryParse(Console.ReadLine(), out selection);
					
					if(regionMenu.IdInList(selection))
					{
						string region = regionMenu.NameById(selection);
						
						Console.WriteLine("You've chosen '{0},' is this correct? Y or N", region);
						
						if(Console.ReadLine().ToUpper().Equals("Y"))
						{
							pathVariable = @"\\oprhp-smb\oprhp_shared\SEQR Clearinghouse\" + region + "\\";
							flag = true;
						}
					}
					else
					{
						Console.WriteLine("You've made an invalid selection, press Enter to try again.");
						Console.ReadLine();
					}
				}
			}
			
			return pathVariable;
		}
		
		//adds project to SEQR workbook
		private static void AddToWorkbook(string parkName, string projectName, bool newProject) {
			string region = albPath.Substring(0, albPath.Length-1);
			region = region.Substring(region.LastIndexOf("\\")+1);
			string smRegion = region.Replace(" ", "");
			string workbook = albPath + smRegion + "SEQRWorkbook.xlsx";
			//workbook = "D:\\CentralSEQRWorkbook.xlsx";
			
			Excel.Application excelApp = new Excel.Application();
			excelApp.Visible = false;
			
			Excel._Workbook wBk = excelApp.Workbooks.Open(workbook,Type.Missing,Type.Missing,
			           Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,
			           Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing);
			
			Excel._Worksheet wSht = (Excel.Worksheet) wBk.Worksheets[1];
			
			int rowNum = 0;
			
			string searchStr = newProject ? "" : projectName;
			int searchCol = newProject ? 1 : 3;
			
			foreach (Excel.Range row in wSht.Rows) {
				Excel.Range rng = (Excel.Range) wSht.Cells[row.Row, searchCol];

				if(rng.Text.ToString().Equals(searchStr))
				{
					rowNum = row.Row;
					break;					
				}
			}
			
			if(newProject) {
				PrincipalContext ctx = new PrincipalContext(ContextType.Domain);
				UserPrincipal user = UserPrincipal.Current;
				string fullName = user.GivenName + " " + user.Surname;
				
				wSht.Cells[rowNum, 1] = region;
				wSht.Cells[rowNum, 2] = parkName;
				wSht.Cells[rowNum, 3] = projectName;
				wSht.Cells[rowNum, 5] = DateTime.Today.Year;
				wSht.Cells[rowNum, 7] = fullName;
			}else{
				wSht.Hyperlinks.Add(wSht.Cells[rowNum, 4], albPath+parkName+"\\"+projectName,Type.Missing,Type.Missing, parkName+"\\"+projectName);
				wSht.Cells[rowNum, 13] = DateTime.Today;
			}
			
			wBk.Save();
			wBk.Close();
			excelApp.Quit();
			excelApp = null;
			GC.Collect();
		}
		
		//edits text in project description ppt
		private static void EditPPT(string parkName, string projectName) {
			string pPath = localPath + parkName + "\\" + projectName + "\\Project Description, Photos, Map\\Project Description.pptx";
			
			if(!File.Exists(pPath))
				return;
			
			PPT.Application pApp = new PPT.Application();
			PPT.Presentation pPres = pApp.Presentations.Open(pPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
			PPT.Slide pSlide = pPres.Slides[1];
			
			foreach(PPT.Shape s in pSlide.Shapes) {
				if(s.Name.Equals("Title 2")) {
					s.TextFrame.TextRange.Delete();
					s.TextFrame.TextRange.InsertBefore(projectName);
				}
				
				if(s.Name.Equals("Text Placeholder 3")) {
					s.TextFrame.TextRange.Delete();
					s.TextFrame.TextRange.InsertBefore(parkName);
				}
			}
			
			pPres.Save();
			pPres.Close();
			pApp.Quit();
			pApp = null;
			GC.Collect();
		}
		
		//edits text in SEQR Checklist
		private static void EditChecklist(string parkName, string projectName) {
			string region = albPath.Substring(0, albPath.Length-1);
			region = region.Substring(region.LastIndexOf("\\")+1);
			string wPath = localPath + parkName + "\\" + projectName + "\\SEQR and Checklist\\SEQRChecklist and Classification Form.docx";
			
			if(!File.Exists(wPath))
				return;
			
			object path = wPath;
			object confirmConversions = Type.Missing;
			object readOnly = false;
			object addToRecentFiles = false;
			object passwordDocument = Type.Missing;
			object passwordTemplate = Type.Missing;
			object revert = Type.Missing;
			object writePasswordDocument = Type.Missing;
			object writePasswordTemplate = Type.Missing;
			object format = Type.Missing;
			object encoding = Type.Missing;
			object visible = false;
			object openAndRepair = false;
			object documentDirection = Type.Missing;
			object noEncodingDialog = true;
			object xmlTransform = Type.Missing;
			
			Word.Application wApp = new Word.Application();

			Word.Document wDoc = wApp.Documents.Open(ref path, ref confirmConversions, ref readOnly, ref addToRecentFiles, 
			          ref passwordDocument, ref passwordTemplate, ref revert, ref writePasswordDocument, ref writePasswordTemplate, 
			          ref format, ref encoding, ref visible, ref openAndRepair, ref documentDirection, ref noEncodingDialog,
			          ref xmlTransform);
			
			object regionIndex = 1;
			object parkIndex = 2;
			object projectIndex = 3;
			
			wDoc.ContentControls[regionIndex].Range.Delete();
			wDoc.ContentControls[regionIndex].Range.InsertBefore(region);
			wDoc.ContentControls[parkIndex].Range.Delete();
			wDoc.ContentControls[parkIndex].Range.InsertBefore(parkName);
			wDoc.ContentControls[projectIndex].Range.Delete();
			wDoc.ContentControls[projectIndex].Range.InsertBefore(projectName);
			
			object saveChanges = true;
			object originalFormat = Type.Missing;
			object routeDocument = Type.Missing;
			
			((Word._Document) wDoc).Close(ref saveChanges, ref originalFormat, ref routeDocument);
			((Word._Application) wApp).Quit();
			wApp = null;
			GC.Collect();
		}
		#endregion
	}
}