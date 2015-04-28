package importador;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ImportarExcel {

	/**
	 * Columnas empleadores: 
	 * 				0.- Sector	
	 * 				1.- Empleador	
	 * 				2.- Domicilio Empleador - (Calle / Numero-Localidad-Provincia-Pais)	
	 * 				3.- CUIT (00-00000000-0) (Empleador)
	 * 				4.- Legajo (Empleado)
	 * 				5.- Nombre y apellido (Empleado)	
	 * 				6.- CUIL (00-00000000-0) (Empleado)	
	 * 				7.- Feriado (Liquida Feriado/No liquida Feriado)
	 */
	public static void main(String[] args) throws BiffException, IOException, SQLException, ClassNotFoundException {

		// Creo la coneccion
	    String myDriver = "org.gjt.mm.mysql.Driver";
	    String myUrl = "jdbc:mysql://localhost/lsc_schema";
	    Class.forName(myDriver);
	    Connection conn = DriverManager.getConnection(myUrl, "ale", "123456");
	    Statement st = conn.createStatement();

		// Cargo el archivo
		File archivo = new File("D:/LSC/workspaces/Importador/Excels/SPAT - AMCA Reporte de Empleadores.xls"); 
		// Lo leo como excel
		Workbook empleadoresExcel = Workbook.getWorkbook(archivo);
		// Hagarro la primer hoja
		Sheet empleadoresSheet = empleadoresExcel.getSheet(0);

		String domicilio;

		// Recorremos las filas (Excluyendo los nombres de columnas)
		for (int fila = 1; fila < empleadoresSheet.getRows(); fila++) {
			
			String razonSocial = empleadoresSheet.getCell(1, fila).getContents();
			String cuit = empleadoresSheet.getCell(3, fila).getContents();

			// Tiene que tener razon social y CUIT por lo menos
			if (razonSocial != "" && cuit != "") {
				ResultSet rs = st.executeQuery("select * from empleador where razon_social = '" + razonSocial + "'");
				
				if (!rs.next()) {
					// DOMICILIO
					domicilio = empleadoresSheet.getCell(2, fila).getContents();
					domicilio = domicilio.split(" -")[0].replace(" /", "");
					
					// CUIT
					cuit = cuit.replaceAll("-", "");
					
					// SUCURSAL
					String id_sucursal = empleadoresSheet.getCell(0, fila).getContents();
					if (id_sucursal.equals("SPAT")) {
						id_sucursal = "1";
					} else {
						//TODO: provisorio, porque no viene informadas las sucursales
						id_sucursal = "2";
					}
					
					// PAGA FERIADO
					String paga_feriado = empleadoresSheet.getCell(7, fila).getContents();
					if (paga_feriado.equals("No liquida Feriado")) {
						paga_feriado = "0";
					} else {
						paga_feriado = "1";
					}

					// Armo el insert
					StringBuilder insert = new StringBuilder("");
					insert.append("INSERT INTO `lsc_schema`.`empleador` (`razon_social`, `domicilio`, `cuit`, `activo`, `id_sucursal`, `paga_feriado`, `Banco`) VALUES (");
					insert.append("'" + razonSocial + "', ");
					insert.append("'" + domicilio + "', ");
					insert.append(cuit + ", ");
					insert.append("1, ");
					insert.append(id_sucursal + ", ");
					insert.append(paga_feriado + ", ");
					insert.append("'Interbanking')");

					// Ejecuto el insert
					st.execute(insert.toString());
				}
			}
		}
		
		// Cierro la coneccion
		st.close();
	}
}


