/*'*************************************************************************************************************************************************
' Script Name			: Connection to Database
' Description			: Used to create a connection with the database and execute the query to retrieve results
' How to Use			:
'-------------------------------------------------------------------------------------------------------------------------------------------------
' Author                    Version          Creation Date         Reviewer Name         Reviewed Date           Comments 
'-------------------------------------------------------------------------------------------------------------------------------------------------
' Sai Kiran Nataraja         v1.0             03-May-2016
'*************************************************************************************************************************************************
 */
package testUtilities;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.Properties;

/**
 * @author saikiran.nataraja
 *
 */

public class ConnectionToDatabase {


	public static Connection connection;

	/**
	 * Function to create a DB2 connection instance and return the instance of it
	 * @author saikiran.nataraja
	 * @param ServerName
	 * @param ServerPortNumber
	 * @param DatabaseToConnect
	 * @param TSOUserID
	 * @param TSOPassword
	 * @returns Connection object used for ExecuteQuery()
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 */
	public Connection openConnection(String ServerName, String ServerPortNumber, String DatabaseToConnect,
			String TSOUserID, String TSOPassword) throws InstantiationException, IllegalAccessException {
		String jdbcClassName = "com.ibm.db2.jcc.DB2Driver"; 
		String url = "jdbc:db2://" + ServerName + ":" + ServerPortNumber + "/" + DatabaseToConnect; // Set URL for data source
		Properties properties = new Properties(); // Create Properties object
		properties.put("user", TSOUserID); // Set user ID for connection
		properties.put("password", TSOPassword); // Set password for connection
		connection = null;
		try {
			// Load class into memory
			Class.forName ( jdbcClassName).newInstance(); 
			// Establish the connection using the IBM Data Server Driver for JDBC
			connection= DriverManager.getConnection (url,properties);                
			/**Commit changes      * false -manually     */			
			connection.setAutoCommit(true);	
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		} finally {
			if (connection != null) {
				//System.out.println("Connected successfully to the database : "+DatabaseToConnect);
			}
		}
		return connection;
	}

	/**
	 * Function to execute query and print the result
	 * @author saikiran.nataraja
	 * @param con
	 * @param QueryToExecute
	 * @param NumberOfRowsOfDataToExtract
	 * @returns Emptyarray if there are NO rows for the query otherwise we will be displayed with NumberOfRowsOfDataToExtract
	 */
	public String[][] ExecuteQueryFromConnection(Connection con,String QueryToExecute,int... NumberOfRowsOfDataToExtract){
		String retVal="";
		int RowsToExtract=0; //Extract Only one row
		PreparedStatement pst,pst1;
		ResultSet rs,rs1;
		boolean RowsFound=false;
		//Number of Columns to Search for
		int row=0,col=1;
		int GetNumberOfColumns=QueryToExecute.substring(6,QueryToExecute.toUpperCase().indexOf("FROM")).split(",").length;
		String QueryToGetRowCount = QueryToExecute.replace(QueryToExecute.substring(6,QueryToExecute.toUpperCase().indexOf("FROM")), " COUNT(*) ");
		if(NumberOfRowsOfDataToExtract.length>0){
			RowsToExtract=NumberOfRowsOfDataToExtract[0]; //Always fetch the number of rows  prescribed in the call
		}else{
			try{
				pst1=con.prepareStatement(QueryToGetRowCount);
				rs1=pst1.executeQuery();
				if(rs1.next()) {
					RowsToExtract=Integer.parseInt(rs1.getString(1));
				}				
				if(RowsToExtract==0){
					RowsToExtract=1; //To Make the array size =1
				}
			}catch(SQLException e){
				e.printStackTrace();
			}
		}
		String[][] GetQueryOutput=new String[RowsToExtract][GetNumberOfColumns];
		try {
			pst=con.prepareStatement(QueryToExecute);
			rs=pst.executeQuery();
			for(;rs.next() && row <RowsToExtract; row=row+1,col=1){
				for(int columnForArray=0;col<=GetNumberOfColumns;col++,columnForArray++){
					retVal=rs.getString(col);
					GetQueryOutput[row][columnForArray]=retVal;
//					System.out.println("Query Outputs are "+GetQueryOutput[row][columnForArray]); //Debug purpose uncomment this line
					RowsFound=true;
				}
			}
			// Close the ResultSet
			rs.close();
			// Close the Statement
			pst.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		if(!RowsFound){
			return Arrays.copyOf(GetQueryOutput, 0); //return empty array
		}else{
			return GetQueryOutput;
		}
	}

	/**
	 * Function to Close the Connection		
	 * @author saikiran.nataraja
	 * @throws SQLException
	 */
	public void closeConnection() throws SQLException{
		try {
			connection.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}

	}

	/*public static void main(String[] args) throws InstantiationException, IllegalAccessException, SQLException {
		//Query  =SELECT   fetch first 1 rows only
		//HOW TO CALL IN OTHER FUNCTIONS
		ConnectionToDatabase contdb=new ConnectionToDatabase();
		Connection connuction=contdb.openConnection("", "50000", "DSN", "USERNAME", "PASSWORD");
		String[][] ResultOfQuery=contdb.ExecuteQueryFromConnection(connuction,"SELECT DISTINCT,5);
		if(ResultOfQuery.length==0){
			System.out.println("Number of records: 0");
		}else{
			System.out.println("Column value is:"+ResultOfQuery);
		}
		contdb.closeConnection();
	}*/

}