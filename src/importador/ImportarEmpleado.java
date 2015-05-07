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

public class ImportarEmpleado {

	/**
	 * 0.- Sector	
	 * 1.- Sucursal	
	 * 2.- Empleador	
	 * 3.- CUIT Empleador	
	 * 4.- Legajo	
	 * 5.- Nombre y apellido	
	 * 6.- CUIT Empleado	
	 * 7.- Fecha de Alta	
	 * 8.- Jubilado	
	 * 9.- Obra Social
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
		Workbook empleadosExcel = Workbook.getWorkbook(archivo);
		// Agarro la primer hoja
		Sheet empleadosSheet = empleadosExcel.getSheet(0);
		
		// Recorremos las filas (Excluyendo los nombres de columnas)
		for (int fila = 1; fila < empleadosSheet.getRows(); fila++) {
			// Recorro los que tienen un empleador
			String razonSocial = empleadosSheet.getCell(2, fila).getContents();
			if (razonSocial != "") {
				String legajo = empleadosSheet.getCell(4, fila).getContents();
				String apellidoNombre = empleadosSheet.getCell(5, fila).getContents();
				String cuil = empleadosSheet.getCell(6, fila).getContents();
				String fechaAlta = empleadosSheet.getCell(7, fila).getContents();
				String jubilado = empleadosSheet.getCell(8, fila).getContents();
				String obraSocial = empleadosSheet.getCell(9, fila).getContents();
				
				// NOMBRE Y APELLIDO
				String[] nombres = apellidoNombre.split(" ");
				String apellido;
				String nombre;
				int cant = 0;
				
				if (nombres.length > 3) {
					apellido = nombres[0] + " " + nombres[1];
					nombre = nombres[2];
					cant = 3;
				} else {
					apellido = nombres[0];
					nombre = nombres[1];
					cant = 2;
				}
				
				for (int i = cant; i < nombres.length; i ++) {
					nombre = nombre + " " + nombres[i];
				}
				
				//CUIL
				cuil = cuil.replaceAll("-", "");
				
				// FECHA ALTA
				String[] fechaSplit = fechaAlta.split("/");
				
				String ano = fechaSplit[2];
				
				if (ano.length() <= 2) {
					if (new Integer(ano).intValue() > 30) {
						ano = "19" + ano;
					} else {
						ano = "20" + ano;
					}
				}
				
				
				String fechaAltaFormat = ano + "-" + fechaSplit[1] + "-" + fechaSplit[0];

				// Empleador
				String buscarEmpleador = "select id from empleador where razon_social = '" + razonSocial + "'";
				ResultSet rs = st.executeQuery(buscarEmpleador);
				long idEmpleador = 0;

		        while (rs.next()) {
		        	idEmpleador = rs.getLong("id");
		        }

		        // SI el empleador viene en 0, registro el error, sino sigo
		        if (idEmpleador != 0) {
		        	// Insert a ejecutar
			        StringBuilder insert = new StringBuilder("");
			        
					// JUBILADO
					if (jubilado.equalsIgnoreCase("SI")) {
						// Armo el insert
						insert.append("INSERT INTO `lsc_schema`.`empleado` (`legajo`, `nombre`, `apellido`, `cuil`, `fecha_alta`, `jubilado`, `id_empleador`) VALUES (");
						insert.append(legajo + ", ");
						insert.append("'" + nombre + "', ");
						insert.append("'" + apellido + "', ");
						insert.append(cuil + ", ");
						insert.append("'" + fechaAltaFormat + "', ");
						insert.append(1 + ", ");
						insert.append(idEmpleador + ")");
						// Ejecuto el insert
						st.execute(insert.toString());

					} else {
						// OBRA SOCIAL
						String buscarObraSocial = "select id from obra_social where descripcion_corta = '" + obraSocial.toUpperCase() + "'";
						rs = st.executeQuery(buscarObraSocial);
						long idObraSocial = 0;
						
				        while (rs.next()) {
				        	idObraSocial = rs.getLong("id");
				        }
				        
				        // SI la obra social viene en 0, registro el error, sino sigo
				        if (idObraSocial != 0) {
				        	// Armo el insert
							insert.append("INSERT INTO `lsc_schema`.`empleado` (`legajo`, `nombre`, `apellido`, `cuil`, `fecha_alta`, `jubilado`, `id_obra_social`, `id_empleador`) VALUES (");
							insert.append(legajo + ", ");
							insert.append("'" + nombre + "', ");
							insert.append("'" + apellido + "', ");
							insert.append(cuil + ", ");
							insert.append("'" + fechaAltaFormat + "', ");
							insert.append(0 + ", ");
							insert.append(idObraSocial + ", ");
							insert.append(idEmpleador + ")");
							// Ejecuto el insert
							st.execute(insert.toString());
				        } else {
				        	// No aparecio la obra social
				        	System.err.println("No encontre la obra social " + obraSocial + " para el empleado " + apellidoNombre);
				        }
					}
					
		        } else {
		        	// No aparecio el empleador
		        	System.err.println("No encontre el empleador " + razonSocial + " para el empleado " + apellidoNombre);
		        }
			}
		}

		// Cierro la coneccion
		st.close();
	}
}