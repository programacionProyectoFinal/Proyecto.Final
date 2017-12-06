
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

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.FileOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import acm.program.ConsoleProgram;

public class Excel extends ConsoleProgram {
	
	public void run()
	{
		/**
		 * *Se inicializan dos archivos, uno para registros limpios y otro para los registros nash
		 */
		File f= new File("../data/FormatolistaEstudiantes.xlsx");
		File fNash= new File("../data/RegistroSalaNashCentro.xlsx");
		/** 
		 * *Si encuentra los dos archivos entonces
		 */
		if (f.exists() && fNash.exists())
		{
			try {
				/** 
				 * *Crea dos archivos de entrada con los archivos anteriormente inicializados
				 */
				FileInputStream archivo = new FileInputStream(f);
				FileInputStream archivoNash = new FileInputStream(fNash);
				/**
				* Se crean dos listas enlazadas de estudiantes, una alojara los registros limpios y otra los registros nash
				* estas se inicializan llamando a los metodos LeerArchivoExcelLimpio y LeerArchivoExcelNash respectivamente
				* los cuales se especifican mas adeltante
				*/
				LinkedList<Estudiante> listaEstudiantes = LeerArchivoExcelLimpio(archivo);
				LinkedList<Estudiante> listaEstudiantesNash = LeerArchivoExcelNash(archivoNash);
				/**
				 * Luego de que las listas fueron inicializadas estas se comparan con el metodo "comparar"
				 * */
				comparar(listaEstudiantes,listaEstudiantesNash);
				
				/** 
				 * Recorre las dos listas e imprime sus registros, descomenta si asi se quiere
				 */
				/**
				for (int i=0; i<listaEstudiantes.size();i++){
					System.out.println(listaEstudiantes.get(i).getId()+" "+listaEstudiantes.get(i).getNombre()+" "+listaEstudiantes.get(i).getApellido());
				}
				System.out.println(" ");
				for (int i=0; i<listaEstudiantesNash.size();i++){
					System.out.println(listaEstudiantesNash.get(i).getId()+" "+listaEstudiantesNash.get(i).getNombre()+" "+listaEstudiantesNash.get(i).getApellido());
				}*/
				/**
				 * Errores posibles durante la ejecuci�n
				 */
			} catch (FileNotFoundException e) {
				println("No se encontro el archivo");
				e.printStackTrace();
			} catch (IOException e) {
				println("No se pudo inicializar el archivo de entrada");
				e.printStackTrace();
			}
			
		}
	}		
	
	/**
	 * Metodo para leer los archivos limpios, retorna una lista enlazada de estudiantes, recibe el archivo de entrada, 
	 * creado en base a un libro xlsx
	 */
	public LinkedList<Estudiante> LeerArchivoExcelLimpio (FileInputStream archivo)  throws IOException{
		/**
		 * Se crea una lista enlazada de estudiantes, que ser� la que retornar�
		 */
		LinkedList<Estudiante> listaEstudiantes = new LinkedList <Estudiante>();
		/**
		 * inicializa un libro XSSFWorkbook usando el archivo que recibe el metodo como parametro
		 */
		XSSFWorkbook libro = new XSSFWorkbook(archivo);
		/**
		 * inicializa una hoja con la primera hoja del libro anteriormente creado
		 */
		XSSFSheet hoja = libro.getSheetAt(0);
		/**
		 * Iterador utilizado para recorrer cada una de las filas en la hoja
		 */
		Iterator<Row> iteradorTupla = hoja.iterator();
		Row tupla;
		/**
		 * este contador se usa para ignorar la primera tupla que contiene los nombres de los campos
		 */
		int contador = 0;
		/**
		 * Se define un formato para que los numeros de los codigos se impriman sin notaci�n cientifica
		 */
		DecimalFormat formatter = new DecimalFormat("#");
		/**
		 * Recorre cada fila de la hoja mientras aun haya filas
		 */
		while(iteradorTupla.hasNext())
		{
			/**
			 * Contiene la informaci�n de cada fila
			 */
			XSSFRow hssfRow = (XSSFRow) iteradorTupla.next() ;
			/**
			 * Se inicializan los tres atributos que tendran el Estudiante, se utiliza el caracter 'x' para 
			 * indicar cuando un atributo no este presente
			 */
			String id = "x";
			String nombre = "x";
			String apellido = "x";
			/**
			 * Solo agrega registros luego de haber leido la fila con la cabecera de cada columna (Documento, Codigo, Nombre, etc)
			 */
			if (contador > 1 &&(hssfRow.getCell(0)!=null) && hssfRow.getCell(2).getStringCellValue() != ""){
				/**
				 * Para la primera celda de la fila "Documento", obtiene su valor numerico, lo representa siguiendo el formato anteriormente
				 * declarado, finalmente obtiene su valor de string y se le asigna al atributo id 
				 */
				if(hssfRow.getCell(0)!=null){
					if(hssfRow.getCell(0).getCellTypeEnum() == CellType.NUMERIC)
					{
					id = String.valueOf(formatter.format(hssfRow.getCell(0).getNumericCellValue()));
					}else {
						id = hssfRow.getCell(0).getStringCellValue();
					}
				}
				/**
				 * Ignora la segunda celda 
				 */	
					
				if(hssfRow.getCell(2)!=null){
					/**
					 * Para la tercera celda, lee su valor como String, luego crea un arreglo teniendo como base ese String separandolo por comas
					 * asi, la primera componente del arreglo tendr� los apellidos y la segunda los nombres
					 */
					String cadena=  hssfRow.getCell(2).getStringCellValue();
					String [] arregloNombre = cadena.split(",");
					/**
					 * Si el arreglo en la parte que contiene los nombre contiene un espacio, significa que tiene mas de un nombre
					 * por lo que se separa por espacios nuevamente esto nos gener� un arreglo,
					 * se toma la segunda posici�n ya que la primera es una posici�n con "", la segunda tiene el primer nombre
					 * y la tercer� en caso de existir tendria el segundo nombre
					 * Se utiliza la funci�n toLowerCase() para no tener problemas con las comparaciones teniendo todos los nombres
					 * en minuscula
					 */
					//Para tener ambos nombres se elimina el if y del else usamos el cuerpo
					if (arregloNombre[1].contains(" ")){
						nombre = arregloNombre[1].split(" ")[1].toLowerCase();
					}
					/**
					 * En caso de que el nombre no tenga espacios simplemente se asume que solo tiene un nombre por lo que se asigna directamente.
					 */
					else{
						nombre = arregloNombre[1];
					}
					/**
					 * De forma analoga se hace con el apellido
					 */
					if (arregloNombre[1].contains(" ")){
						apellido = arregloNombre[0].split(" ")[0].toLowerCase();
					}else{
						apellido = arregloNombre[0];
					}
				}
				
			}
			/**
			 * Condici�n que asegura que no se agregar� un estudiante con la primera fila (Documento, Codigo, Nombre)
			 */
			if (contador!=0){
				listaEstudiantes.add(new Estudiante(id, apellido, nombre));
			}
			if ((hssfRow.getCell(0)!=null)){
				contador++;
			}
		}
		/**
		 * Se cierra el archivo leido y se retorna la lista de estudiantes
		 */
		libro.close();
		return listaEstudiantes;
		
	}
	/**
	 * Metodo para leer archivos nash, retorna una lista enlazada de estudiantes con los registros leidos, recibe el archivo de entrada, 
	 * creado en base a un libro xlsx
	 */
	public LinkedList<Estudiante> LeerArchivoExcelNash (FileInputStream archivo)  throws IOException{
		/**
		 * El metodo funciona de forma muy parecida al de leer archivos limpios, solo cambia el orden y la forma en que lee las celdas
		 */
		LinkedList<Estudiante> listaEstudiantes = new LinkedList <Estudiante>();
		XSSFWorkbook libro = new XSSFWorkbook(archivo);
		XSSFSheet hoja = libro.getSheetAt(0);
		Iterator<Row> iteradorTupla = hoja.iterator();
		Row tupla;
		//este contador se usa para ignorar la primera tupla que contiene los nombres de los campos
		int contador=0;
		DecimalFormat formatter = new DecimalFormat("#");
		while(iteradorTupla.hasNext())
		{
			XSSFRow hssfRow = (XSSFRow) iteradorTupla.next() ;
			String id = "x";
			String nombre = "x";
			String apellido = "x";
			if (contador > 1){
				/**
				 * Por cada fila se obtiene el documento, el apellido y el nombre de las celdas 0,1 y 2 respectivamente.
				 * Respecto al nombre y el apellido nuevamente se separan por espacios, se toma el primer nombre
				 * y el primer apellido, para no tener problemas con la comparaci�n estos se guardan en minuscula
				 */
				if(hssfRow.getCell(0)!=null){
					id = String.valueOf(formatter.format(hssfRow.getCell(0).getNumericCellValue()));
					/**
					 * Si el valor es 0, significa que no encontro un registro en esta posici�n, por lo que se le asigna 'x'
					 */
					if(hssfRow.getCell(0).getNumericCellValue()==0){
						id = "x";
					}
				}
				
				if(hssfRow.getCell(1)!=null){
					
					apellido = hssfRow.getCell(1).getStringCellValue().split(" ")[0].toLowerCase();	
					/**
					 * Si el valor es vacio, significa que no encontro un registro en esta posici�n, por lo que se le asigna 'x'
					 */
					if(apellido.equals("")){
						apellido = "x";
					}
				}
					
				if(hssfRow.getCell(2)!=null){			
					nombre = hssfRow.getCell(2).getStringCellValue().split(" ")[0].toLowerCase();
					/**
					 * Si el valor es vacio, significa que no encontro un registro en esta posici�n, por lo que se le asigna 'x'
					 */
					if(nombre.equals("")){
						nombre = "x";
					}
				}
			
			}
			/**
			 * A�ade el registro a la lista
			 */
			if (contador!=0){
				listaEstudiantes.add(new Estudiante(id, apellido, nombre));
			}
			if ((hssfRow.getCell(0)!=null)){
				contador++;
			}	
		}
		/**
		 * Cierra el libro y retorna la lista con los estudiantes
		 */
		libro.close();
		return listaEstudiantes;
		
	}
	
	/**
	 * Metodo para Comparar la lista de estudiantes que tienen los registros limpios y la lista con los estudiantes a los que les falta
	 * informaci�n, recibe estas dos listas y no tiene retorno
	 */
	public void comparar (LinkedList<Estudiante> listaEstudiantes, LinkedList<Estudiante> listaEstudiantesNash){
		/**
		 * Crea dos libros y dos hojas nuevas, un libro y una hoja para los registros que se encuentran completamente en ambas listas
		 * y otro par para un archivo de excel que contendr� las alertas
		 */
		XSSFWorkbook libro =  new XSSFWorkbook();
		XSSFSheet hoja = libro.createSheet("1");
		XSSFWorkbook libroAlert =  new XSSFWorkbook();
		XSSFSheet hojaAlert = libroAlert.createSheet("1");
		/**
		 * Contadores de fila para que al escribir las filas estas no se sobreescriban
		 */
		int contador = 1;
		int contadorAlert = 1;
		/**
		 * Recorre cada estudiante en la lista de estudiantes a los que le puede faltar o no informaci�n
		 */
		for (int i=0;i<listaEstudiantesNash.size();i++){
			boolean condIgualdad = false;
			/**
			 * Compar� cada estudiante en la lista nash con los de la lista de registros limpios
			 */
			for (int j=0;j<listaEstudiantes.size();j++){
				boolean idIgual = false;
				boolean nombreIgual = false;
				boolean apellidoIgual = false;
				/**
				 * Si encuentra un registro cuya informaci�n sea igual en ambas listas, reasigna las condiciones de igualdad 
				 * de cada atributo como verdaderas
				 */
				if (listaEstudiantesNash.get(i).getId().equals(listaEstudiantes.get(j).getId())) {
					idIgual=true;
				}
				if (listaEstudiantesNash.get(i).getNombre().equals(listaEstudiantes.get(j).getNombre())) {
					nombreIgual=true;
				}
				if (listaEstudiantesNash.get(i).getApellido().equals(listaEstudiantes.get(j).getApellido())) {
					apellidoIgual=true;
				}
				/**
				 * Si los tres atributos son iguales se asigna la condifici�n de igualdad entre ambos registros como verdadera
				 */
				if (idIgual && nombreIgual && apellidoIgual){
					condIgualdad = true;
				}
			}
			/**
			 * Crea una fila en la primera posici�n de la hoja que contendr� los registros de salida, 
			 * Les asigna el valor de "Documento", "Apellido" y "Nombre" a las tres primeras celdas de la fila
			 * esto para los titulos
			 */
			try
			{

			XSSFRow fila = hoja.createRow(0);
			XSSFCell celda0 = fila.createCell(0);
			XSSFCell celda1 = fila.createCell(1);
			XSSFCell celda2 = fila.createCell(2);
			celda0.setCellValue("Documento");
			celda1.setCellValue("Apellido");
			celda2.setCellValue("Nombre");
				/**
				 * Crea un archivo llamado "salida.xlsx"
				 * escribe en este la fila anteriormente creada
				 * y cierra el archivo
				 */
			FileOutputStream output = new FileOutputStream("../data/salida.xlsx"); 
			libro.write(output); 
			/**
			 * De forma parecida se hace con el archivo "Alertas.xlsx" que tendr� las alertas encontradas durante la comparac�on
			 * Se le asigna un titulo de columna "Alertas" y se escribe en el archivo.
			 */
			fila = hojaAlert.createRow(0);
			celda0 = fila.createCell(0);
			celda0.setCellValue("Alertas");
			FileOutputStream output1 = new FileOutputStream("../data/alertas.xlsx"); 
			libroAlert.write(output1); 
				 			
			/**
			 * Si encuentra algun atributo en los archivos nash sin informaci�n, crea una fila en la posici�n que indique el contadorAlert,
			 * y en la primera celda de esa fila escribe que hay que revisar la fila con el registro incompleto,
			 * tambien imprime esta alerta.
			 */		
			if((listaEstudiantesNash.get(i).getId().equals("x")) || (listaEstudiantesNash.get(i).getApellido().equals("x")) || (listaEstudiantesNash.get(i).getNombre().equals("x"))){
				fila = hojaAlert.createRow(contadorAlert);
				contadorAlert = contadorAlert+1;
				celda0 = fila.createCell(0);
				celda0.setCellValue("revisar la fila "+(i+2));
				println("revisar la fila "+(i+2));
				libroAlert.write(output1);  
			}
			/**
			 * Si encontro dos registros iguales en ambas listas, crea una fila en la hoja de registros segun indique el contador,
			 * crea 3 celdas en esta fila, y las inicializa con el documento, el apellido y el nombre respectivamente.
			 * escribe la fila en el archio "salida.xlsx"
			 */	
			else if (condIgualdad){
				fila = hoja.createRow(contador);
				contador++;
				celda0 = fila.createCell(0);
				celda1 = fila.createCell(1);
				celda2 = fila.createCell(2);
				celda0.setCellValue(listaEstudiantesNash.get(i).getId());
				celda1.setCellValue(listaEstudiantesNash.get(i).getApellido());
				celda2.setCellValue(listaEstudiantesNash.get(i).getNombre());
				libro.write(output); 
					 
			/**
			 * Si ninguna de las anteriores condiciones se dieron significa que el estudiante simplemente no existe en el listado
			 * de registros limpios, nuevamente se crea una fila, en la primera celda se escribe que el estudiante de la respectiva fila
			 * no existe, se escribe esta celda en el archivo "alertas-xlsx" y se cierra	
			 */
			}else if (!condIgualdad){
				fila = hojaAlert.createRow(contadorAlert);
				contadorAlert++;
				celda0 = fila.createCell(0);
				celda0.setCellValue("el estudiante de la fila "+(i+2)+" no existe");
				println("el estudiantede la fila "+(i+2)+" no existe");
				libroAlert.write(output);
			}
			output.close(); 
			output1.close();
			}
			catch (Exception e) {
				println("El archivo no pudo escribirse");
				}
			}
		
		}
}


"Clase Estudiantes"
package excel_apachePOI;

public class Estudiante {
	//cada estudiante tiene una id, un apellido y un nombre
	private String id;
	private String apellido;
	private String nombre;
	
	//constructor de la clase estudiante, recibe un double que representa la id, un String para apellido y un String para nombre.
	public Estudiante(String id, String apellido, String nombre) {
		this.id = id;
		this.apellido = apellido;
		this.nombre = nombre;
		
	}

	//metodos para asegurar el encapsulamiento, con los get se trae el valor del atributo, los set se usan para asignarlo
	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
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
	

}
