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

public class ImportarEmpleado {

	public static void main(String[] args) throws BiffException, IOException, SQLException, ClassNotFoundException {
		// Creo la coneccion
	    String myDriver = "org.gjt.mm.mysql.Driver";
	    String myUrl = "jdbc:mysql://localhost/lsc_schema";
	    Class.forName(myDriver);
	    Connection conn = DriverManager.getConnection(myUrl, "ale", "123456");
	    Statement st = conn.createStatement();

		// Cargo el archivo
		File archivo = new File("D:/LSC/workspaces/Importador/Excels/"); 
		// Lo leo como excel
		Workbook empleadosExcel = Workbook.getWorkbook(archivo);
		// Agarro la primer hoja
		Sheet empleadosSheet = empleadosExcel.getSheet(0);
		
		// Recorremos las filas (Excluyendo los nombres de columnas)
		for (int fila = 1; fila < empleadosSheet.getRows(); fila++) {
			
		}

		// Cierro la coneccion
		st.close();
	}
}
