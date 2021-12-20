package info;

import java.util.*;

import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class TableInfo
{
	public enum eOrientation { Vertical, Horizontal };
	
	protected int _NumberOfLines = 0 ;
	protected int _NumberOfColumns =0 ;
	protected eOrientation _Orientation = eOrientation.Vertical;
	protected int _HeaderDepth = 0;
	
	protected XWPFTable _Table;
	protected ArrayList<Integer> _LinesPerRow = null;

	/* ==========================================================================
	 * 
	 * 
	 *  ----- Constructor
	 * 
	 * ==========================================================================
	 */
	public TableInfo( XWPFTable pTable )
	{
		_LinesPerRow = new ArrayList<Integer>();
		_Table = pTable;
		
		_NumberOfLines = findNumberOfLines();
		_NumberOfColumns = findNumberOfColumns();
		_Orientation = findOrientation();
	}
	
	/* ==========================================================================
	 * 
	 * 
	 *  ----- Accessors
	 * 
	 * ==========================================================================
	 */
	
	public int getNumberOfLines() 
	{
		return _NumberOfLines;
	}
	
	public int getNumberOfColumns()
	{
		return( _NumberOfColumns );
	}
	
	public eOrientation getOrientation()
	{
		return( _Orientation );
	}
	
	public int getHeaderDepth()
	{
		return( _HeaderDepth );
	}

	public int getLinesForRow( int pRow )
	{
		return( _LinesPerRow.get( pRow ) );
	}
	
	
	
	/*
	 *  =======================================================
	 * 
	 *  Get Number Of Lines
	 *  Try to define the max number of lines in a row.
	 *  Based on the number of inner paragraphs.
	 * 
	 *  =======================================================
	 */
	protected int findNumberOfLines()
	{
		List<XWPFTableRow> rows;
		XWPFTableRow row;
		XWPFTableCell cell;
		int rowLines;
		int numberOfLines;
		
		//
		// ---- Init
		//
		numberOfLines = 0;
		rows = _Table.getRows();
		
		
		//
		// ---- Loop over each cell to find the max in each row. then add.
		//
		for( int rowNo = 0; rowNo < rows.size(); rowNo++ )
		{
			row = rows.get(rowNo);
			rowLines = 0;
			
			for( int columnNo = 0; columnNo < row.getTableCells().size(); columnNo++) 
			{
				cell = row.getTableCells().get(columnNo);
				rowLines = Math.max( rowLines, CellInfo.getNumberOfLines( cell ) );
			}
			_LinesPerRow.add( rowLines );
			
			numberOfLines += rowLines;
		}
		
		return( numberOfLines );
	}
	
	
	
	/*
	 *  =======================================================
	 * 
	 *  Get Number Of Columns
	 *  Try to define the max number of columns in a row.
	 *  Based on the number of cells
	 * 
	 *  =======================================================
	 */
	protected int findNumberOfColumns()
	{
		XWPFTableRow row;
		int numberOfColumns = 0;
		
		row = _Table.getRows().get(0);
		for( XWPFTableCell cell : row.getTableCells() )
		{
			numberOfColumns += CellInfo.getNumberOfColumns(cell);
		}
		
		return( numberOfColumns );
	}
	
	
	/*
	 *  ====================================================================
	 * 
	 *  ---- Find Orientation
	 * 
	 *  ====================================================================
	 */
	public eOrientation findOrientation()
	{
		eOrientation direction = eOrientation.Vertical;
		int rowNo;
		boolean rowIsHeader;
		boolean hasStopped;
		
		//
		// ---- Loop over all rows
		//
		hasStopped = false;
		for( rowNo = 0; rowNo < _Table.getRows().size(); rowNo++ )
		{
			rowIsHeader = isHeaderCellSet( _Table.getRows().get(rowNo).getTableCells() );
			
			//
			// ---- if first row is special then array is horizontal
			//
			if( (rowNo == 0)  && rowIsHeader )
			{
				direction = eOrientation.Horizontal;
			}
			else
			{
				direction = eOrientation.Vertical;
			}
			
			//
			//  ----- Keep track of the last special row
			//
			if( rowIsHeader && !hasStopped )
			{
				_HeaderDepth = rowNo + 1;
			}
			else
			{
				hasStopped = true;
			}
			
		}
				
		return( direction );
	}
	
	/*
	 * ================================================================= 
	 * 
	 * ---- Check if an entire set is special, making it eligible for
	 *      header candidate.
	 * 
	 * ==================================================================
	 */
	public static boolean isHeaderCellSet( List<XWPFTableCell> list )
	{
		boolean isHeader = true;
		boolean isHeaderCell;
		CellInfo info;
		
		for( XWPFTableCell cell : list )
		{
			info = new CellInfo( cell );
			
			isHeaderCell = info.isHeaderCell();
			if( !isHeaderCell  )
			{
				isHeader = false;
			}
		}
		
		return( isHeader );
	}
	
	
	
	

}
