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
				String[] nombres = apellidoNombre.split("");
				String apellido;
				String nombre;
				
				if (nombres.length > 2) {
					apellido = nombres[0] + " " + nombres[1];
					nombre = nombres[2] + " " + nombres[3];
				} else {
					apellido = nombres[0];
					nombre = nombres[1];
				}
				
				//CUIL
				cuil = cuil.replaceAll("-", "");
				
				// FECHA ALTA
				String[] fechaSplit = fechaAlta.split("/");
				String fechaAltaFormat = fechaSplit[2] + "-" + fechaSplit[1] + "-" + fechaSplit[0];

				// Empleador
				String buscarEmpleador = "select id from empleador where razon_social = '" + razonSocial + "'";
				ResultSet rs = st.executeQuery(buscarEmpleador);
				long idEmpleador = 0;
				
		        while (rs.next()) {
		        	idEmpleador = rs.getLong("id");
		        }

		        // Insert a ejecutar
		        StringBuilder insert = new StringBuilder("");
		        
				// JUBILADO
				if (jubilado.equalsIgnoreCase("NO")) {
					// Armo el insert
					insert.append("INSERT INTO `lsc_schema`.`empleado` (`legajo`, `nombre`, `apellido`, `cuil`, `fecha_alta`, `jubilado`, `id_empleador`) VALUES (");
					insert.append(legajo + ", ");
					insert.append("'" + nombre + "', ");
					insert.append("'" + apellido + "', ");
					insert.append(cuil + ", ");
					insert.append("'" + fechaAltaFormat + "', ");
					insert.append(0 + ", ");
					insert.append(idEmpleador + ")");

				} else {
					// OBRA SOCIAL
					String buscarObraSocial = "select id from obra_social where descripcion_corta = '" + obraSocial + "'";
					rs = st.executeQuery(buscarObraSocial);
					long idObraSocial = 0;
					
			        while (rs.next()) {
			        	idEmpleador = rs.getLong("id");
			        }
					
					// Armo el insert
					insert.append("INSERT INTO `lsc_schema`.`empleado` (`legajo`, `nombre`, `apellido`, `cuil`, `fecha_alta`, `jubilado`, `id_obra_social`, `id_empleador`) VALUES (");
					insert.append(legajo + ", ");
					insert.append("'" + nombre + "', ");
					insert.append("'" + apellido + "', ");
					insert.append(cuil + ", ");
					insert.append("'" + fechaAltaFormat + "', ");
					insert.append(1 + ", ");
					insert.append(idObraSocial + ", ");
					insert.append(idEmpleador + ")");					
				}

				// Ejecuto el insert
				st.execute(insert.toString());
			}
		}

		// Cierro la coneccion
		st.close();
	}
}

/*
 * 
NO ESTA: 
104504 Obra Social Personal de la Industria Cinematografi
O. S. Boxeadores Agremiados de la Rep. Arg.
O.S. DEL PERSONAL LADRILLERO

REVISAR:
O.S.ASOC. MUTUAL OBREROS CATOLICOS PADRE FEDERICO GROTE
O.S. P/PERSONAL DE OBRAS Y SERV.SANITARIOS
O.S.ACCION SOCIAL DE EMPRESARIOS
O.S.ALF.REPOST. PIZZEROS,ETC
O.S.COMERCIO
O.S.C.E.A.R.B.A
O.S.DE LOS MEDICOS DE LA CIUDAD DE BS.AS
O.S.CONDUCTORES,TITULARES DE TAXIS DE LA CABA
O.S.DE LA UNION OBRERA METALURGICA DE LA REP. ARG.
O.S.DE SERENOS DE BUQUES
O.S.DE LOS TRABAJADORES DE LA IND. DEL GAS
O.S.DE SUPERVISORES DE LA IND. METALMECANICA DE LA REP ARG
O.S.DEL PERS. RURAL Y ESTIBADORES DE LA REP. ARG.
O.S.DEL PERSONAL AERONAUTICO
O.S.DEL PERS. DE EDIF. DE RENTA Y HORIZONTAL DE LA C.A.B.A.
O.S.ELECTRICISTAS NAVALES
O.S.DEL PERSONAL DE MAESTRANZA
O.S.DEL PERSONAL DE INSTALACIONES SANITARIAS
O.S.DEL PERSONAL DEL AUTOMOVIL CLUB ARGENTINO
O.S.DEL PERSONAL JERARQUICO DE AGUA Y ENERGIA
O.S.DEL PERSONAL DE SEGURIDAD IND E INVEST, PRIV.
O.S.DEL PERSONAL RURAL Y ESTIBADORES DE LA REP ARG
O.S.PARA LA ACTIVIDAD DOCENTE
O.S.EMPLEADOS TEXTILES Y AFINES
O.S.PELUQUERIA, ESTETICA Y AFINES
O.S.EMPRESARIOS, PROFESIONALES Y MONOTRIBUTISTAS
O.S.GNC,GARAGES,PLAYAS DE ESTAC.Y LAVADEROS AUTOMATICOS
O.S.FEDERACION DE CAMARAS Y CENTROS COMERCIALES ZONALES
O.S.PASTELEROS, CONFITEROS,PIZZEROS Y ALFAJOREROS DE LA REP.
O.S.EMPLEADOS Y PERSONAL JURIDICO DEL NEUMATICO ARG.GOOD YEA
OS DE ARBITROS
OS BANCARIA ARGENTINA
O.S.PERS.ACTIV.DEL TURF
OS COND. TRANSP. COLECTIVO
O.S.PERS.SOC.AUTORES Y AFINES
OBRA SOCIAL DE SERENOS DE BUQUES
obra social empleados de comercio
OBRA SOCIAL DEL PERSONAL DE SANIDAD
O.S.PERSONAL JERARQUICO DE AGUA Y ENERGIA
OS CAP. DE ULTRAMAR Y OFIC. MARINA MERCANTE
OBRA SOCIAL DEL PERSONAL SUPERIOR MERCEDES BENZ
OS EMPL DE COMERCIO
OS ORGS CONTROL EXTERNO
OS DEL PERSONAL DE AERONAVEGACION
OS DE COMIS. NAVALES
OS PERS DE LA CONSTRUCCION
OS PERSONAL GRAFICO
OS DE CAPATACES Y ESTIBADORES PORTUARIOS
OS DE CHOFERES DE CAMIONES
OS PORTUARIOS ARG
OS DEL PERS. JERARQUICO DE LA REPUBLICA ARGENTINA
OS EMPL. DE DESPACHANTE DE ADUANA
OS DE RELOJEROS Y JOYEROS
OS PERS. IND. DEL PLASTICO
OS DEL P. DE LA IND. DEL FIBROCEMENTO
OS SANIDAD
OS PESONAL DE PUBLICIDAD
OS TELECOMUNICACIONES
OS UNION DEL PERS CIVIL DE LA NACION
OSMMEDT
OS TRAB DE PRENSA BA
OS. PERS. IND. MOLINERA
OS TRABAJADORES DEL GAS
OS VIAJANTES VENDEDORES DE LA REP.ARG.
OSADEF - FARMACIAS
OSPEDYH
OSUTHGRA
*/
