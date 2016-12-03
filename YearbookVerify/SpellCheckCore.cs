using Microsoft.Office.Interop.Excel;
using NetSpell.SpellChecker;
using NetSpell.SpellChecker.Dictionary;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace YearbookVerify {
	/// <summary>
	/// Models a spell checking engine for the storing of a name dictionary,
	/// and the processing of text to see if it is in the dictionary, as well
	/// as if it has the proper pairings.
	/// </summary>
	class SpellCheckCore : IDisposable {
		//constants
		private const int MAX_SUGGESTIONS = 2;
		private const string NAME_PAIRINGS_INPUT_FILE = "/pairings.xlsx";
		private const string NAME_PAIRINGS_STORAGE_FILE = "/storedPairings.nms";
		private const string VALID_NAMES_FILE_READ = "validNames.dic"; //note: no / since api for spellchecker is a bit funky
		private const string VALID_NAMES_FILE_WRITE = "/validNames.dic";
		private const int STARTING_EXCEL_INPUT_ROW = 3;
		private const int FIRST_NAME_EXCEL_INPUT_COL = 2;
		private const int LAST_NAME_EXCEL_INPUT_COL = 3;
		//members
		private Spelling spellCheck;
		private char lineDelim;
		private List<Name> registeredPairs;
		//constructors
		public SpellCheckCore(char lineDelim) {
			this.lineDelim = lineDelim;
			spellCheck = new Spelling();
			spellCheck.SuggestionMode = Spelling.SuggestionEnum.PhoneticNearMiss;
			spellCheck.Dictionary = LoadValidNames();

			registeredPairs = LoadNamePairs();
		}
		//methods
		/// <summary>
		/// Attempts to read the stored name pairs from the storage file.
		/// On error, it alerts the user that the name paring database hasn't been built.
		/// </summary>
		private List<Name> LoadNamePairs() {
			try {
				List<Name> names = new List<Name>();
				string[] lines = File.ReadAllLines(Environment.CurrentDirectory + NAME_PAIRINGS_STORAGE_FILE);
				//Parse all string lines into Name objects
				return Array.ConvertAll(lines, Name.Parse).ToList();
			}
			catch {
				//name storage not found, database hasn't been built yet
				MessageBox.Show("No name pairings have been loaded into the spell checker. "
					+ "Please Reload the Name Database after entering valid name pairings into  " + NAME_PAIRINGS_INPUT_FILE + ".", 
					"Error Reading Pairings File", MessageBoxButton.OK);
			}
			return new List<Name>();
		}
		/// <summary>
		/// Saves all name parings to file
		/// </summary>
		private void StoreNamePairs() {
			try {
				File.WriteAllText(Environment.CurrentDirectory + NAME_PAIRINGS_STORAGE_FILE,
					string.Join(Environment.NewLine, registeredPairs));
			}
			catch (Exception e) {
				MessageBox.Show("An error occured attempting to save valid name parings to file.\n\nError:\t" 
					+ e.Message, "Error Writing Names File", MessageBoxButton.OK);
			}
		}
		/// <summary>
		/// Loads a <seealso cref="WordDictionary"/> from file to use
		/// as the database of all possible names
		/// </summary>
		private WordDictionary LoadValidNames() {
			if (!File.Exists(Environment.CurrentDirectory + VALID_NAMES_FILE_WRITE)) {
				StoreValidNames(new List<string>());
			}
			WordDictionary wd = new WordDictionary();
			wd.DictionaryFile = VALID_NAMES_FILE_READ;
			return wd;
		}
		/// <summary>
		/// Stores all valid names given to file
		/// </summary>
		private void StoreValidNames(List<string> valids) {
			try {
				File.WriteAllText(Environment.CurrentDirectory + VALID_NAMES_FILE_WRITE,
					Properties.Resources.dicBase + string.Join(Environment.NewLine, valids));
			}
			catch (Exception e) {
				MessageBox.Show("An error occured trying to save the valid name pairings. \n\nError:\t" 
					+ e.Message, "Error Writing Dictionary File", MessageBoxButton.OK);
			}
		}
		/// <summary>
		/// Finds all the unique, singular name strings from each registered
		/// first-name last-name pair
		/// </summary>
		/// <returns>the sorted list of unique names</returns>
		private List<string> NamesToUniqueWords() {
			List<string> unique = new List<string>();
			foreach (Name n in registeredPairs) {
				if (!unique.Contains<string>(n.First))
					unique.Add(n.First);
				if (!unique.Contains<string>(n.Last))
					unique.Add(n.Last);
			}
			unique.Sort();
			return unique;
		} 
		/// <summary>
		/// Provides a few suggestions for spelling mistakes to the given word strings
		/// </summary>
		/// <param name="words">as many strings to provide suggestions for as necessary</param>
		/// <returns>a string listing the suggestions</returns>
		private string FindSomeSuggestions(params string[] words) {
			List<string> suggs = new List<string>();
			foreach (string word in words) {
				//find suggestions
				spellCheck.Suggest(word);
				//add a max of 2 suggestions/word
				for (int i = 0; i < spellCheck.Suggestions.Count && i < MAX_SUGGESTIONS; i++) {
					suggs.Add(Name.FixupCapitals(spellCheck.Suggestions[i] as string));
				}
				
			}
			//join all suggestions together
			string sugg = string.Join(" or ", suggs);
			return sugg.Length > 0 ? " Did you mean: " + sugg + "?" : "";
		}
		/// <summary>
		/// Checks the spelling of all name pairs in a text string.
		/// The names are separated by the line delimiter set in the constructor.
		/// Mispelled names are noted and given suggestions for fixes
		/// </summary>
		/// <returns>A <seealso cref="SpellingResult"/> that contains the editted original lines and the errors reported</returns>
		public SpellingResult CheckSpelling(string text) {
			string[] lines = text.Split(lineDelim);
			//lines the user put in, to be editted for captitalization
			string[] userLines = new string[lines.Length];
			//lines that mark errors in spelling and name pairing
			string[] markedLines = new string[lines.Length];
			for (int i = 0; i < lines.Length; i++) {
				if (lines[i].Length == 0) {
					//skip blank lines
					userLines[i] = "";
					markedLines[i] = "";
					continue;
				}
				//parse a name from the line
				Name parsed = Name.Parse(lines[i]);
				if (parsed != null) {
					//correct any user capitalization mistakes
					userLines[i] = parsed.ToString();
					//check the spelling of each name
					bool sFirst = spellCheck.TestWord(parsed.First);
					bool sLast = spellCheck.TestWord(parsed.Last);
					//both correct
					if(sFirst && sLast) {
						//see if the name is a registered name pairing
						Name matchedName = registeredPairs.Find(n => n.IsSame(parsed));
						//name pair not registered, but names spelled correctly
						if(matchedName == null) {
							markedLines[i] = "Error: \"" + parsed + "\" is not registered name pair.";
						}
						else {
							//all good!
							markedLines[i] = "";
						}
					}
					else if (!sFirst && !sLast) {
						//both mispelled
						markedLines[i] = "Error: Both \"" + parsed.First + "\" and \""
										+ parsed.Last + "\" are unregistered names. " 
										+ FindSomeSuggestions(parsed.First, parsed.Last);
					}
					else if (!sFirst) {
						//first is bad
						markedLines[i] = "Error: \"" + parsed.First + "\" is not a registered first name. " 
							+ FindSomeSuggestions(parsed.First);
					}
					else {
						//last is bad
						markedLines[i] = "Error: \"" + parsed.Last + "\" is not a registered last name. "
							+ FindSomeSuggestions(parsed.Last);
					}
				}
				else {
					//no name could be parsed, too long or too short
					throw new SpellCheckException("Error with input: \"" + lines[i]
						+ "\". Each name must be from 2 to 4 words. If this is an issue, contact me to fix it.");
				}
			}
			return new SpellingResult(string.Join("\n", userLines), string.Join("\n", markedLines));
		}
		/// <summary>
		/// Reloads the spellchecker registered pair databases,
		/// then the dictionary of all valid names.
		/// </summary>
		public void ReloadData() {
			registeredPairs = CompileNamesFromExcel();
			StoreNamePairs();
			StoreValidNames(NamesToUniqueWords());
			spellCheck.Dictionary.Dispose();
			spellCheck.Dictionary = LoadValidNames();
		}
		/// <summary>
		/// Reads the input worksheet and parses valid first/last name pairings.
		/// </summary>
		private List<Name> CompileNamesFromExcel() {
			try {
				List<Name> compNames = new List<Name>();
				//new hidden excel app
				Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
				app.Visible = false;
				//open the input.xlsx workbook, first worksheet
				Workbook wb = app.Workbooks.Open(Environment.CurrentDirectory + NAME_PAIRINGS_INPUT_FILE);
				Worksheet ws = wb.Worksheets[1];

				//parse names
				string first, last;
				int i = STARTING_EXCEL_INPUT_ROW;
				bool done = false;
				do {
					first = (string)ws.Cells[i, FIRST_NAME_EXCEL_INPUT_COL].Value;
					last = (string)ws.Cells[i, LAST_NAME_EXCEL_INPUT_COL].Value;
					done = first == null || last == null || first.Length == 0 || last.Length == 0;
					if (!done) {
						compNames.Add(new Name(first, last));
						i++;
					}
				}
				while (!done);
				//release resources
				wb.Close(false);
				app.Quit();
				ReleaseInterop(ws);
				ReleaseInterop(wb);
				ReleaseInterop(app);
				return compNames;
			}
			catch (Exception e) {
				MessageBox.Show("An error occured when trying to read the input spreadsheet. Make sure the file " 
					+ NAME_PAIRINGS_INPUT_FILE + " is located in the directory" +
					"with this program and that it is formatted properly. The first first name should be in cell B3 and the first last name should be in cell" +
					" C3.\n\nError:\t" + e.Message, "Error reading spreadsheet", MessageBoxButton.OK);
			}
			return new List<Name>();
		}
		/// <summary>
		/// Release the resources associated with the given interop object
		/// </summary>
		private void ReleaseInterop(object obj) {
			try {
				System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
				obj = null;
				//good
			}
			catch (Exception ex) {
				obj = null;
				MessageBox.Show("An error has occured (Can't release memory): " + ex.Message);
				//error
			}
			finally {
				GC.Collect();
			}
		}
		public void Dispose() {
			spellCheck.Dispose();
		}
	}
	/// <summary>
	/// Models an error in checking spelling
	/// </summary>
	class SpellCheckException : Exception {
		public SpellCheckException(string message) : base(message) { }
	}
	/// <summary>
	/// Packages up the results of a spelling operation
	/// </summary>
	class SpellingResult {
		//properties
		internal string UserLines { get; private set; }
		internal string MarkedLines { get; private set; }
		//constructors
		internal SpellingResult(string userLines, string markedLines) {
			UserLines = userLines;
			MarkedLines = markedLines;
		}
	}
}
