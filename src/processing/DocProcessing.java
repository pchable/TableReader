package processing;

import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;

public class DocProcessing
{
	final static String cCAPTION_STYLENAME = "Lgende";
	final static int cCAPTION_MAX_LENGTH = 40;
	final static int cMAX_TITLE_DISTANCE = 5;
	final static String cTABLE_TITLE= "Tableau";
	
	protected TableProcessing _TableProcessing;

	/*
	 *  =======================================================
	 * 
	 *  Constructor
	 * 
	 *  =======================================================
	 */
	
	public DocProcessing()
	{
		_TableProcessing = new TableProcessing();
	}

	/*
	 *  =======================================================
	 * 
	 *  Find Tables in the document.
	 * 
	 *  =======================================================
	 */	
	public void findTables( String pWordDocName )
	{
		InputStream file;
		XWPFDocument document = null;
		XWPFTable table;
		int tableNo = 0;
		int elementNo = 0;
		String caption, targetFileName;
		

		try
		{
			//
			// ---- Open Docx
			//
			file = new FileInputStream( pWordDocName );  
			document = new XWPFDocument( file );
			
			document.getBodyElements().size();
			
			//
			// ---- Loop Over All Elements
			//
			for( elementNo = 0; elementNo < document.getBodyElements().size(); elementNo++ )
			{
				IBodyElement element = document.getBodyElements().get( elementNo );
				
				switch( element.getElementType() )
				{
					case TABLE:		table = (XWPFTable)element;
									caption = getCaption( document, elementNo, tableNo );
									targetFileName = getTableFileName( pWordDocName, caption );
									_TableProcessing.ProcessTable( table, caption, targetFileName );
									tableNo++;
									break;
									
					default:		break;
				}
			}
			
			//
			// ---- and close
			//
			document.close();
		}
		catch( Exception pException )
		{
			pException.printStackTrace();
		}  
	}
	
	/*
	 *  =======================================================
	 * 
	 *  Retrieve Table Caption: either next paragraph or make it up
	 * 
	 *  =======================================================
	 */
	protected String getCaption( XWPFDocument pDocument, int pElementNo, int pTableNo )
	{
		XWPFParagraph paragraph = null;
		String caption;
		String style;
		String text;
		
		//
		// ---- Default Value
		//
		caption = Integer.toString(pTableNo);
		
		//
		// ---- Try to retrieve table name if available within a reasonable distance
		//      ie within the next paragraphs
		//
		try
		{
			for( int attempt= 0; attempt < cMAX_TITLE_DISTANCE; attempt++ )
			{
				paragraph = (XWPFParagraph)(pDocument.getBodyElements().get( pElementNo + attempt + 1 ));
				style = paragraph.getStyle();
				text = paragraph.getText();
					
				//
				// ---- Sanity check
				//
				if( (style == null) || (text == null) )
				{
					continue;
				}
				
				//
				// ----- If the following paragraph contains "Tableau" and is of the "Legend" type, may be a caption.
				//
				if( style.equals(cCAPTION_STYLENAME) && text.contains(cTABLE_TITLE)) 
				{
					caption = text;
					break;
				}
			}
		}
		catch( Exception pException )
		{
			System.out.println("Could not retrieve caption for table: " + pTableNo );
		}
		
		
		//
		// ---- In any case return something.
		//
		return( truncate( caption, cCAPTION_MAX_LENGTH ) );
	}
	
	
	/*
	 *  =======================================================
	 * 
	 *  Remove inadequate characters and limit caption size
	 * 
	 *  =======================================================
	 */
	protected String truncate( String pSource, int pMaxLength )
	{
		String result ="";
		char letter;
		int length;
		
		length = Math.min( pSource.length(), pMaxLength);
		
		for( int index = 0; index < length; index++ )
		{
			letter = pSource.charAt(index);
			switch( letter )
			{
				case ' ':
				case ':':	
				case '\t':	
				case '\n':	
				case '\f': 
				case '\r': 	letter = '_';
							break;
				
				default: 	break;
			}
						
			result += letter;
		}
		
		return( result );
	}
	
	/*
	 *  =======================================================
	 * 
	 *  Create Table FileName
	 * 
	 *  =======================================================
	 */
	public String getTableFileName( String pWordDocumentName, String pCaption )
	{
		String fileName;
		Path path;
		
		path = Paths.get(pWordDocumentName);
		fileName = ".\\" + path.getFileName();
		fileName = fileName.substring(0, fileName.length() - 5);  // remove the 5 letter extension ".docx"
		fileName += "_CAPTION_" + pCaption + ".txt";
		
		return( fileName );
	}

}
