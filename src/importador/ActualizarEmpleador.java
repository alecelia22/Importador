package importador;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ActualizarEmpleador {
	
	/**
	 * 0.- Sector
	 * 1.- Sucursal
	 * 2.- Empleador
	 * 
	 */
	public static void main(String[] args) throws BiffException, IOException, SQLException, ClassNotFoundException {

		// Creo la coneccion
	    String myDriver = "org.gjt.mm.mysql.Driver";
	    String myUrl = "jdbc:mysql://localhost/lsc_schema";
	    Class.forName(myDriver);
	    Connection conn = DriverManager.getConnection(myUrl, "ale", "123456");
	    Statement st = conn.createStatement();

		// Cargo el archivo
		File archivo = new File("D:/LSC/workspaces/Importador/Excels/SPAT - AMCA Reporte de Empleados y Empleados.xls"); 
		// Lo leo como excel
		Workbook empleadoresExcel = Workbook.getWorkbook(archivo);
		// Agarro la primer hoja
		Sheet empleadoresSheet = empleadoresExcel.getSheet(0);

		// Recorremos las filas (Excluyendo los nombres de columnas)
		for (int fila = 1; fila < empleadoresSheet.getRows(); fila++) {
			
			String razonSocial = empleadoresSheet.getCell(2, fila).getContents();
			
			String sucursal = empleadoresSheet.getCell(1, fila).getContents();
			
			if (razonSocial != null && !razonSocial.equals("") && sucursal != null && !sucursal.equals("")) {
			
				int numeroSucursal = 2;
	
				if (sucursal.equalsIgnoreCase("Autodromo")) {
					numeroSucursal = 2;
				} else if (sucursal.equalsIgnoreCase("Belgrano")) {
					numeroSucursal = 4;
				} else if (sucursal.equalsIgnoreCase("Caballito")) {
					numeroSucursal = 5;
				} else if (sucursal.equalsIgnoreCase("Casa central")) {
					numeroSucursal = 6;
				} else if (sucursal.equalsIgnoreCase("Chacarita")) {
					numeroSucursal = 7;
				} else if (sucursal.equalsIgnoreCase("Correo Interno")) {
					numeroSucursal = 8;
				} else if (sucursal.equalsIgnoreCase("Lima")) {
					numeroSucursal = 9;
				} else if (sucursal.equalsIgnoreCase("Mosconi")) {
					numeroSucursal = 10;
				} else if (sucursal.equalsIgnoreCase("Premetro")) {
					numeroSucursal = 11;
				} else if (sucursal.equalsIgnoreCase("Quilmes")) {
					numeroSucursal = 12;
				} else if (sucursal.equalsIgnoreCase("San Justo")) {
					numeroSucursal = 13;
				}
	
				String sql = "update empleador set id_sucursal = " + numeroSucursal + " where razon_social = '" + razonSocial + "';";
				
				// Ejecuto el insert
				st.execute(sql);
			}
		}
		
		// Cierro la coneccion
		st.close();
	}
}



