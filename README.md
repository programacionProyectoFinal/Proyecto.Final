//* En el avance del proyecto final de programación estamos trabajando en un código cuyo propósito temporal es crear filas y columnas de 
celdas tales que podamos despúes crear metodos para comparar esta información creada. También creamos un método que permite separar nombres 
y apellidos, ya que la hoja limpia tiene estos dos atributos separados unicamente por una coma. Hasta el momento presentamos varios problemas,
como lo es la lectura de hojas de Excel ya creadas y al momento de crear nuestras propias hojas el codigo nos rechaza los datos creados 
anteriormente.
Los proximos avances del codigo serán los siguientes:
  1. se solucionará los errores que tenemos hasta el día de hoy con el código.
  2. seremos capaces de leer hojas de Excel ya creadas y poder modificarlas a gusto.
  3. crearemos los metodos que llevará el programa principal, los cuales serán capaces de recorrer la lista con ayuda de un for y
  comparar los datos entre hojas de Excel 
  4. generaremos la alerta que indicará cuando algún dato no exista y nos muestre las sugerencias de cómo corregir dicho error.
  
Por fin podemos leer datos de una hoja excel para extraerlos en forma de ArrayList y asociarcelos al objeto de la clase Estudiante. Con esto ya obtendremos acceso a la información de cada estudiante en los archivos de la sala Nash


Clase para crear hojas de Excel:
package excel_apachePOI;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import com.sun.org.apache.xpath.internal.operations.Variable;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import java.io.FileNotFoundException;
import java.io.IOException;

import acm.program.ConsoleProgram;

public class Excel extends ConsoleProgram {
	public void run() {
		XSSFWorkbook workbook = new XSSFWorkbook(); // En el workbook se trabaja todo
		Sheet sh1 = workbook.createSheet();
		Row row = sh1.createRow(0);// 0 es la primera
		//Cell cell = row.createCell(4); // row = fila
		Cell documento = row.createCell(0);
		
		/*for (int col = 0; col < 5;col++)
		{
			variable  = sh1.createRow(col).createCell(0);
			variable.setCellValue("Hola mundo");
		}
		Cell variable1 = sh1.createRow(2).createCell(0);
		variable1.setCellValue("chao");
		//variable.setCellValue(variable1.getStringCellValue());
		
		if (variable1.getStringCellValue()!= variable.getStringCellValue() ) {
			variable1.setCellValue("que vaina tan jodida");
			
			
		}*/
		
		
		
		
		//Cell cell2  = sh1.createRow(1).createCell(5);
		//cell2.setCellValue(cell.getStringCellValue()); // Hay que especificar en el metodo el tipo de dato de la celda
		
		//Cell cell_error = sh1.createRow(3).createCell(9);
		//cell_error.setCellValue("ola mundo");
		
		
		//if (cell_error.getStringCellValue() != cell.getStringCellValue()) {
			//cell_error.setCellValue(cell.getStringCellValue());
		//}
		
		int contador = 6;
		
		String [] documentos1 = { "CEDULA",
				"1007402218",
				"1020809175",
				"1085345980",
				"99080311008",
				"99090408140",
				"99110500390"};
		for (int k = 0 ; k < 7;k++)
		{
			documento = sh1.createRow(contador).createCell(3);
			documento.setCellValue(documentos1[k]);
			contador ++;
		}  
		
		Cell codigo = row.createCell(0);
		int contador1 = 14;
		String [] codigo1 = {"codigo","171132184","161655117","171655083","171655016","171655100","171655125"};
		for (int f = 0 ; f < 7; f ++)
		{
			codigo = sh1.createRow(contador1).createCell(4);
			codigo.setCellValue(codigo1[f]);
			contador1 ++;
		}
		
		Cell nombre = sh1.createRow(6).createCell(4);
		int contador2 = 6;
		String [] nombre1 = {"APELLIDOS Y NOMBRES",
				"LEGUIZAMON GARCIA, LAURA DANIELA",
				"GUTIERREZ PARADA, OSCAR JAVIER",
				"CAICEDO RAMIREZ, NICOLAS GIOVANNY",
				"CARDOZO GONZALEZ, JUAN DAVID",
				"CASTILLO HERRERA, CRISTIAN ALEXANDER",
				"DIAZ LIZARAZO, GABRIELA",};
		for (int m = 0; m < 7;m ++)
		{
			nombre = sh1.createRow(contador2).createCell(4);
			nombre.setCellValue(nombre1[f]);
			contador2 ++;
		}
		try {
			//Workbook wb = WorkbookFactory.create(new FileInputStream("C:\\\\Users\\prestamour\\Downloads\\RegistroNashCentro.xlsx"));
			FileOutputStream output = new FileOutputStream("E:\\\\ProyectoEnsayo_ApachePOI\\data\\Registro1.xlsx"); //para salvarlo en el hardrive se crea el FileOutputStream
			workbook.write(output); //escribir ese output en el FileOutputStream
			output.close(); // cerrar output por seguridad (como FileReader)
		}catch (Exception e) {
			println("El archivo no pudo leerse");
		}
	}

}

Clase para separar nombres y apellidos

package excel_apachePOI;
import acm.program.*;

public class SepararNombreApellido extends ConsoleProgram {
	public static final String nombCompleto = "Ruiz Ortiz, Juan Camilo";
	public void run() {
		separarNombresApellidos(nombCompleto);
		
	}
	public String[] separarNombresApellidos(String NombAp) {
		String dato = "";
		int pos = 0;
		String [] nomApeSep = new String[NombAp.length()];
		
		for(int i =0; i < nomApeSep.length; i++) {
			if(NombAp.substring(i, i+1) == " "){
				nomApeSep [i] = NombAp.substring(i, i+1);
			}
			if(NombAp.substring(i,i+1) == "," && NombAp.substring(i+1,i+2) == " ") {
				nomApeSep [i] = NombAp.substring(i, i+1);	
			}
		return nomApeSep;
			}
			
		}
		
      }
      
      
      nueva clase estudiantes
      package excel_apachePOI;
Hemos modificado la clase Estudiantes para que reciba un ArrayList, desconponga sus datos y extraiga los datos que sean necesarios. De momento ese ArrayList sólo cuenta con los datos: Documento, apellido, nombre (suponemos también que los datos extraídos siempre vienen en ese orden para no tener incoherencias).
import acm.program.*;

package miEstudiante;

import java.util.ArrayList;

public class Estudiante {
	private double documento;
	private String apellido;
	private String nombre;
	private ArrayList datos;
	// Suponiendo que el orden del ArrayList es [documento, apellido, nombre]
	public Estudiante(ArrayList misDatos) {
		datos = misDatos;
		this.documento = (double) misDatos.get(0);
		this.apellido = (String) misDatos.get(1);
		this.nombre = (String) misDatos.get(2);
		
	}

	public double getDocumento() {
		return documento;
	}

	public void setDocumento(double documento) {
		this.documento = documento;
	}

	public String getApellido() {
		return apellido;
	}

	public void setApellido(String apellido) {
		this.apellido = apellido;
	}

	public String getNombre() {
		return nombre;
	}

	public void setNombre(String nombre) {
		this.nombre = nombre;
	}
	public String toString() {
		return this.getDocumento() +"\n" + this.getApellido()+ "\n"+ this.getNombre(); 
	}
	

}

hemos mejoreado el codigo que lee y escribe los nuevos datos, aprendimos a usar el StringTokenizer ya que en los datos limpios el apellido y nombre se encuentran en la misma celda y para compararlos necesitamos que esten en celdas diferentes

	package excel_apachePOI;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sun.org.apache.xpath.internal.operations.Variable;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;

import java.io.FileOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import acm.program.ConsoleProgram;

public class Excel extends ConsoleProgram {
	
	
	public void run()
	{
		File f = new File("C:\\\\Users\\prestamour\\Documents\\ejemplo1.xlsx");
		if (f.exists())
		{
			leerArchivosDeExcel(f);
		}
	}
	
	public void escribirArchivoDeExcel(String valor0,String valor1,String valor3,int j)
	{
		HSSFWorkbook libro = new HSSFWorkbook();
		HSSFSheet hoja = libro.createSheet();
		HSSFRow fila = hoja.createRow(j);
		HSSFCell celda0 = fila.createCell(0);
		celda0.setCellValue(valor0);
		HSSFCell celda1 = fila.createCell(1);
		celda1.setCellValue(valor1);
		HSSFCell celda2 = fila.createCell(2);
		celda2.setCellValue(valor3);
		
		try
		{
			FileOutputStream output = new FileOutputStream("E:\\\\ProyectoEnsayo_ApachePOI\\data\\Registro1.xlsx"); 
			libro.write(output); 
			output.close(); 
		}catch (Exception e) {
			println("El archivo no pudo leerse");
		}
	}
	
	
	
	private void obtener(List cellDataList)
	{
		
		for (int i = 0;i < cellDataList.size();i ++)
		{
			List cellTempList = (List) cellDataList.get(i);
			//for(int j = 0;j < 3;j++)
			//{
				
				XSSFCell xssfcell0 = (XSSFCell) cellTempList.get(0);
				XSSFCell xssfcell1 = (XSSFCell) cellTempList.get(1);
				XSSFCell xssfcell3 = (XSSFCell) cellTempList.get(3);
				String cellvalue0 = xssfcell0.toString();
				String cellvalue1 = xssfcell1.toString();
				String cellvalue3 = xssfcell3.toString();
				escribirArchivoDeExcel(cellvalue0,cellvalue1,cellvalue3,i);
				/*print(cellvalue0 +"\t");
				print(cellvalue1 +"\t");
				print(cellvalue3 +"\t");
				println(" ");*/
			//}
			
		
		}
		
	}
	
	
	public void leerArchivosDeExcel(File fileName) {
		List TodosEstudiantes = new ArrayList();
		try {
			FileInputStream fileInputStream = new FileInputStream(fileName);
			XSSFWorkbook workBook = new XSSFWorkbook(fileInputStream);
			XSSFSheet hssfSheet = workBook.getSheetAt(0);
			Iterator rowIterator = hssfSheet.rowIterator();
			while(rowIterator.hasNext())
			{
				XSSFRow hssfRow = (XSSFRow) rowIterator.next();
				Iterator iterator = hssfRow.cellIterator();
				List estudiante = new ArrayList();
				while (iterator.hasNext())
				{
					XSSFCell hssfCell = (XSSFCell) iterator.next();
					estudiante.add(hssfCell);			
				}
				TodosEstudiantes.add(estudiante);
				
			}
		}catch (Exception e) {
			println("El archivo no pudo leerse");
		}
		obtener(TodosEstudiantes);
		println(TodosEstudiantes.get(0));
		
		
	}

	


}	
		





