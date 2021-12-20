package processing;

import java.awt.Point;
import java.io.PrintWriter;


import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import info.CellInfo;
import info.TableInfo;


public class TableProcessing 
{
	final static String cTITLE_SEPARATOR = " ";
	XWPFTable _Table;
	TableInfo _TableInfo;
	String[][] _Data;

	public TableProcessing() 
	{	
		_Table = null;
		_TableInfo = null;
		_Data = null;
	}
	
		

	
	/*
	 *  =======================================================
	 * 
	 *  Process Single Table
	 * 
	 *  =======================================================
	 */
	public void ProcessTable( XWPFTable pTable, String pTableCaption, String pTargetFile )
	{
		PrintWriter writer;
		
		//
		// ---- Init
		//
		_Table = pTable;
		_TableInfo = new TableInfo( pTable );
		
		System.out.println("\n\n\nProcessing table: " + pTableCaption );
		System.out.println("==================================================");
		System.out.println("\tLines: " +  _TableInfo.getNumberOfLines() );
		System.out.println("\tColumns: " + _TableInfo.getNumberOfColumns());
		System.out.println("\tOrientation: " + _TableInfo.getOrientation()  );
		System.out.println("\tDepth: " + _TableInfo.getHeaderDepth() + "\n" );
		
		//
		// ---- Skip single line table
		//
		if( _TableInfo.getNumberOfLines() <2 )
		{
			return;
		}
		
		//
		// ----- Generate File for Serialized Table
		//
		try
		{
			writer = new PrintWriter( pTargetFile );
			writer.print("TABLE BEGIN\n\n");
			
			serializeTable( writer );
			
			writer.print("\nTABLE END");
			writer.close();
		}
		catch( Exception pException )
		{
			System.out.println("Error found during table processing ! - " + pException.getMessage() );
		}
		
	}
	
	/*
	 *  =======================================================
	 * 
	 *  Transform Array to Single Line
	 * 
	 *  =======================================================
	 */
	public void serializeTable( PrintWriter pWriter )
	{
		int line, column;
		String title, value, header;

		//
		// ---- Retrieve data from Table
		//
		buildDataArray();
		ShowData();
			
		
		
		//
		// ---- Loop over each row (skip header hence start from 1)
		//
		for( line = _TableInfo.getHeaderDepth(); line < _TableInfo.getNumberOfLines(); line++  )
		{
		
			pWriter.print("\tROW BEGIN\n\t\t");
			
			//
			// ---- Write for each cell the value next to the column header.
			//
			for( column = 0; column < _TableInfo.getNumberOfColumns(); column++  )
			{
				//
				// ---- Build-up the column header on the possible multiple first lines
				//     happens in some presentation with the main section & sub sections.
				//
				title = "";
				for( int depth = 0; depth < _TableInfo.getHeaderDepth(); depth++ )
				{
					//
					// ----- From second header line, add separator
					//
					if( depth != 0 )
					{
						title += cTITLE_SEPARATOR;
					}
					
					header = _Data[depth][ column ];
					if( header == null )
					{
						header = "element";
					}
					
					title += header;  
				}
				
				//
				// ---- Find value & write.
				//
				value  = _Data[ line ][column ];				
				pWriter.print( "["+ title +"]" + "=" + value  +";");
			}
			
			pWriter.print("\n\tROW END\n");
		}
		
	}
	
	
	
	
	
	/*
	 *  =======================================================
	 * 
	 *  ----- Build an array filled with all cell values
	 * 
	 *  =======================================================
	 */
	public void buildDataArray()
	{
		XWPFTableRow row;
		XWPFTableCell cell;
		
		Point cursor, offset;
		int rowNo, blocNo;
		

		//
		// ---- init
		//
		_Data = new String[ _TableInfo.getNumberOfLines() ][ _TableInfo.getNumberOfColumns() ];
		cursor = new Point( 0, 0);
		offset = new Point( 0, 0 );
		
		
		//
		// ---- Loop over all the table content
		//
		for( rowNo = 0; rowNo < _Table.getNumberOfRows(); rowNo++ )
		{
			row = _Table.getRow(rowNo);
			
			for( blocNo = 0; blocNo < row.getTableCells().size(); blocNo ++)
			{
				cell = row.getCell(blocNo);
				
				offset = setCellData( cell, cursor, rowNo );
				cursor.y += offset.y;
			}		
			cursor.x += _TableInfo.getLinesForRow(rowNo);
			
			//
			// ---- reset for next line.
			//
			cursor.y = 0;
		}
		
	}
	
	
	/*
	 *  =======================================================
	 * 
	 * 
	 *  ----- Try for each cell to retrieve the content
	 * 
	 * 
	 *  =======================================================
	 */
	public Point setCellData( XWPFTableCell pCell, Point pCursor, int pRowNo )
	{
		Point offset;
		CellInfo info;
		int hspan, col;
		int line, linesInRow;
		String firstLine;
		XWPFTableRow row;
		
		//
		// ---- Init
		//
		offset = new Point( 0, 0 );
		info = new CellInfo( pCell );
		hspan = info.getNumberOfColumns();
		linesInRow = _TableInfo.getLinesForRow(pRowNo);
		row = _Table.getRow( pRowNo );
		
		//
		// ---- Spread the topline content to all cells horizontally merged
		//      Note: if not merged, then cell span is 1
		//
		firstLine = info.getContent(0);
		for( col = 0; col < hspan; col++ )
		{
			_Data[ pCursor.x][ pCursor.y + col ] = firstLine;
		}
		offset.y = col;
		
		//
		// ---- Multiple Lines must spread over several lines except header
		//
		if( TableInfo.isHeaderCellSet(row.getTableCells() ) )
		{
			_Data[ pCursor.x ][ pCursor.y ] = info.normalize( pCell.getText() );
		}
		else
		{
			if( info.hasMultipleLines() )
			{
				for( line = 0; line < info.getNumberOfLines(); line++ )
				{
					_Data[ pCursor.x + line ][ pCursor.y ] = info.getContent(line);
				}
			}
		}
		
		//
		// ---- If single line but the row has multiple, then duplicate
		//
		if( ! info.hasMultipleLines() &&  (_TableInfo.getLinesForRow(pRowNo) > 1) )
		{
			for( line = 1; line < linesInRow; line++ )
			{
				_Data[ pCursor.x + line ][ pCursor.y ] = firstLine;
			}
			
		}
			
				
		//
		// ---- Vertical Fusion
		// ----- copy the value of the top to all merged cells below.
		//
		try
		{
			if( info.isVMerged() && info.getContent(0).isEmpty() )
			{
				
				for( line = 0; line < linesInRow; line++ )
				{
					_Data[ pCursor.x + line ][ pCursor.y ] = _Data[ pCursor.x - 1][ pCursor.y ];
				}
			}
		}
		catch( Exception pException )
		{
			System.out.println("Failed whilst trying to duplicate line above: ");
			System.out.println("\tLine: " + pCursor.x );
			System.out.println("\tColumn: " + pCursor.y );
		}

		return( offset );
	}
	
	
	/* ===========================================================================
	 * 
	 * 
	 *  ---- Dump Procedure.
	 * 
	 * 
	 * ==========================================================================
	 */
	public void ShowData( )
	{
		String format = " %-25s|";
		
		for( int line = 0; line < _TableInfo.getNumberOfLines(); line++ )
		{
			System.out.print("\n\t");
			for( int column = 0; column < _TableInfo.getNumberOfColumns(); column++ )
			{
				System.out.printf( format, _Data[ line ][ column ]);
			}
			
		}
	}
	
}
