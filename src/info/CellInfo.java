package info;

import java.util.ArrayList;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;


// Test de connexion avec GitHib

public class CellInfo
{
	final static String cAUTO = "auto";
	final static String cWHITE = "FFFFFF";
	final static String cPowerOfTen = ".10";
	
	XWPFTableCell _Cell;
	protected boolean _isHeader = false;
	protected ArrayList<String> _Content = null;
	
	
	/* ========================================================================
	 * 
	 *  Constructor which analyse the real content of a cell
	 * 
	 * 
	 * ========================================================================
	 */
	public CellInfo( XWPFTableCell pCell )
	{
		_Cell = pCell;
		
		findContent();
		_isHeader = findIsHeader();

		
	}
	
	/* ========================================================================
	 * 
	 *  
	 *  Check if Cell is somehow special
	 * 
	 * ========================================================================
	 */
	public boolean findIsHeader()
	{
		boolean isHeader = false;
		boolean hasHeaderLook = false;
		boolean hasHeaderValue = false;
		
		
		hasHeaderLook = hasHeaderLook();
		hasHeaderValue = hasHeaderValue();
		
		if( hasHeaderLook & hasHeaderValue )
		{
			isHeader = true;
		}
		
		return( isHeader );			
	}
	
	/* ========================================================================
	 * 
	 *  This routine converts the color code to avoid that "Auto" or "White" are
	 *  considered as highlighting color codes.
	 * 
	 * ========================================================================
	 */
	protected String normalizeColor( String pColor )
	{
		String color = null;
		
		//
		// ---- Sanity Check
		//
		if( pColor == null )
		{
			return( null );
		}
		
		switch( pColor )
		{
			case cAUTO:
			case cWHITE:	color = null;
							break;
							
			default:		color = pColor;
							break;
		}
		
		return( color );
	}
	
	
	/* ========================================================================
	 * 
	 *  Check if the cell has a style eligilible to header
	 * 
	 * =========================================================================
	 */
	public boolean hasHeaderLook()
	{
		boolean hasHeaderLook = false;
		boolean hasColoredBackground = false;
		boolean hasHeaderDecoration = false;
		boolean isBold = false; 
		String color;
		CTTcPr cellProperties;
		CTTcBorders borders;
		
		
		//
		// ---- Check if first row has no decoration in which case it could be header.
		//
		try
		{
			cellProperties = _Cell.getCTTc().getTcPr();
			borders = cellProperties.getTcBorders();
			
			if( borders.getTop().getVal() == STBorder.NIL || 
				borders.getLeft().getVal() == STBorder.NIL )
			{
				hasHeaderDecoration = true;
			}
		}
		catch( Exception pException )
		{
			// well... no border specific ! 
		}

		
		//
		// ---- If colored, probably a header
		//
		color = normalizeColor( _Cell.getColor() );
		if( color != null )
		{
			hasColoredBackground = true;
		}
		
		//
		// ---- If bold, probably a header
		//
		for( XWPFRun run : _Cell.getParagraphs().get(0).getRuns() )
		{
			if( run.isBold() )
			{
				isBold = true;
			}
		}
		
		hasHeaderLook = hasColoredBackground | isBold | hasHeaderDecoration ;
		
		return( hasHeaderLook );
	}
	
	public boolean isHeaderCell()
	{
		return( _isHeader );
	}
	
	/* ========================================================================
	 * 
	 *  Retrieve Cell possible texts
	 *  
	 * 
	 * ========================================================================
	 */
	
	public void findContent()
	{
		
		//
		// ---- Init
		//
		_Content = new ArrayList<String>();
		
		//
		// ---- Split cell content by line
		//
		for( XWPFParagraph paragraph : _Cell.getParagraphs() )
		{
			_Content.add( normalize( paragraph.getText() ) );
		}

	}
	
		
	public String getContent( int pIndex )
	{
		if( pIndex > _Content.size() )
		{
			return( null );
		}
		else
		{
			return( _Content.get( pIndex ) );
		}
	}
	
	/* ========================================================================
	 * 
	 *  Manage Vertical Merge / Multiple lines per cell.
	 * 
	 * ========================================================================
	 */
	public boolean isVMerged()
	{
		boolean isVMerged = false;
		
		if( _Cell.getCTTc().getTcPr().isSetVMerge() )
		{
			isVMerged = true;
		}
		
		return( isVMerged );
	}
	
	public int getNumberOfLines()
	{
		return( getNumberOfLines( _Cell ) );
	}
	
	/*
	 *  =================================================================
	 * 
	 * 
	 * ---- Header lines only have one line... ever.
	 * 
	 * ==================================================================
	 */
	static protected int getNumberOfLines( XWPFTableCell pCell )
	{
		XWPFTableRow row;
		int numberOfLines = 1;
		
		row = pCell.getTableRow();
		
		
		if( ! TableInfo.isHeaderCellSet( row.getTableCells()  ) )
		{
			numberOfLines = pCell.getParagraphs().size();
		}
		
		return( numberOfLines );
	}
	
	public boolean hasMultipleLines()
	{
		return( (getNumberOfLines() > 1) ); 
	}

	
	/* ================================================================
	 * 
	 * 
	 *  Manage Horizontal Merge
	 * 
	 * ================================================================
	 */
	
	public boolean isHMerged()
	{
		boolean isHMerged = false;
		
		if( _Cell.getCTTc().getTcPr().isSetGridSpan())
		{
			isHMerged = true;
		}
		
		return( isHMerged );
	}
	
	public int getNumberOfColumns()
	{
		return( getNumberOfColumns( _Cell ) );
	}
	
	static public int getNumberOfColumns( XWPFTableCell pCell )
	{
		int numberOfColumns = 1;
		
		if( pCell.getCTTc().getTcPr().isSetGridSpan() )
		{
			numberOfColumns = pCell.getCTTc().getTcPr().getGridSpan().getVal().intValue();
		}
		
		return( numberOfColumns );
	
	}
	
	/*
	 *  =======================================================
	 * 
	 *  ----- Normalize Strings (remove special characters )
	 * 
	 *  =======================================================
	 */
	public String normalize( String pSource )
	{
		String result;
		char letter;
		
		result = "";
		
		for( int index = 0; index < pSource.length(); index++ )
		{
			letter = pSource.charAt(index);
			
			switch( letter )
			{
				case '\t':
				case '\n':
				case '\r':
				case '\f':	letter = ' ';
							break;
							
				default:	break;
			}
			
			result += letter;
		}
		
		return( result );
	}
	
	
	/*
	 * ================================================================= 
	 * 
	 * ---- A cell could be a header cell if
	 *      - it contains no value 
	 *        or
	 *      - if the value it contains is NOT a number
	 * 
	 * ==================================================================
	 */
	public boolean hasHeaderValue()
	{
		
		String content;
		
		content = _Cell.getText();
		
		//
		// ----- Check is the cell contains something
		//
		boolean hasNoText = false;
		if( content.equals("") )
		{
			hasNoText = true;
		}
		
		//
		// ---- If contains something like "3,4.105" then may be a number
		//
		boolean isPowerOfTenNumber = false ;
		if( content.contains( cPowerOfTen ) )
		{
			isPowerOfTenNumber = true;
		}
		
		//
		// ---- Check if value converts to number.
		//      if no exception triggered, then this is true
		//      otherwise false
		//
		//       /!\ Use of french format converter with comma for decimal separation.
		//
		boolean isNotANumber = false;
		String candidateNumber = content.replace(",", ".");
		try
		{
			Double.valueOf(candidateNumber);
		}
		catch( Exception pException )
		{
			isNotANumber = true;
		}
		
		//
		// ----- Consistency
		//
		if( isPowerOfTenNumber )
		{
			isNotANumber = false; // because it is !!!
		}

		
		return( hasNoText | isNotANumber  );
	}
	

}
