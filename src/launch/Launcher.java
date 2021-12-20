package launch;

import processing.DocProcessing;

public class Launcher 
{

	public Launcher() 
	{
	}

	/* ==================================================================================
	 * 
	 * 
	 * ----- Main
	 * 
	 * ==================================================================================
	 */
	public static void main(String[] args) 
	{

		DocProcessing doc = new DocProcessing();
		
		//
		// ---- Start Up
		System.out.println("Running Table Extraction...");
		
		//
		// ---- Checking input
		//
		if( args.length != 1 )
		{
			System.out.println("Le nombre d'arguments est incorrect ! ");
		}
		else
		{
			System.out.println("Processing file: \n\t" + args[0] );
			doc.findTables( args[0] );
		}
	

	}
	
	

}
