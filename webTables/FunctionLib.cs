/*
 * Created by Ranorex
 * User: tom
 * Date: 12/12/2018
 * Time: 21:19
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

// add this for file writing
using System.IO;


namespace webTables
{
    /// <summary>
    /// Ranorex user code collection. A collection is used to publish user code methods to the user code library.
    /// </summary>
    [UserCodeCollection]
    public class FunctionLib
    {
        /// <summary>
        /// Takes a web table, exports all the cells to a csv file
        /// returns: nothing
        /// parameters: web table repo item, path to csv output file
        /// dependencies: system.io for file writing
        /// </summary>
        [UserCodeMethod]
        public static void WebToExcel(Ranorex.Adapter htmlTable, String path)
        {
        	// the false parameter means not to append but overwrite the file if it exists:
        	TextWriter textWriter = new StreamWriter(path, false, System.Text.Encoding.UTF8);
        	
        	// Get the Web Table Rows (have to convert the repo item to web table tag)
        	IList<Ranorex.TrTag> IRows = htmlTable.As<Ranorex.TableTag>().FindDescendants<Ranorex.TrTag>();
			// Loop through each row and get the cells
        	foreach (Ranorex.TrTag theRow in IRows) {
        		// Get the Cells in Each Row
        		IList<Ranorex.TdTag> ICells = theRow.FindDescendants<Ranorex.TdTag>();
				// loop through each cell in the row
				String StrCSV = "";
        		foreach (Ranorex.TdTag theCell in ICells) {
        			StrCSV += theCell.GetInnerHtml() + ",";
        		}
				textWriter.WriteLine(StrCSV);     		
        	}        	
			textWriter.Close();
        }
        
        
        
        /// <summary>
        /// Compares two text files and reports to results any differences
        /// returns: TRUE or FALSE if files are same
        /// Parameters: paths to the two files
        /// dependencies: system.io
        /// </summary>
        [UserCodeMethod]
        public static bool CompareFiles(String path1, String path2)
        {
        	using (StreamReader f1 = new StreamReader(path1))
			using (StreamReader f2 = new StreamReader(path2)) {
			
			    String differences = "";
			
			    int lineNumber = 0;
			
			    while (!f1.EndOfStream) {
			        if (f2.EndOfStream) {
			           differences += "Differing number of lines - f2 has less." + Environment.NewLine;
			           break;
			        }
			
			        lineNumber++;
			        var line1 = f1.ReadLine();
			        var line2 = f2.ReadLine();
			
			        if (line1 != line2) {
			        	differences += string.Format("Line {0} differs. File 1: {1}, File 2: {2}", lineNumber, line1, line2) + Environment.NewLine;
			        }
			    }
			
			    if (!f2.EndOfStream) {
			         differences += "Differing number of lines - f1 has less." + Environment.NewLine;;
			    }
			    if (differences != ""){
			    	Report.Info(differences);
			    	return false;
			    }
			    else{
			    	return true;
			    }
			}
        }
        
        
    }
}
