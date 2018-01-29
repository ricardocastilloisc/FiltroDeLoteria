import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.InputMismatchException;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	static int matriz[][];
	static int fila;

	public static void main(String[] args) throws Exception {

	}

	public void Ejecutar() throws IOException {
		//
		// An excel file name. You can create a file name with a full
		// path information.
		//
		String filename = "excel/bd.xls";
		//
		// Create an ArrayList to store the data read from excel sheet.
		//
		List sheetData = new ArrayList();
		FileInputStream fis = null;
		try {
			//
			// Create a FileInputStream that will be use to read the
			// excel file.
			//
			fis = new FileInputStream(filename);
			//
			// Create an excel workbook from the file system.
			//
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			//
			// Get the first sheet on the workbook.
			//
			HSSFSheet sheet = workbook.getSheetAt(0);
			//
			// When we have a sheet object in hand we can iterator on
			// each sheet's rows and on each row's cells. We store the
			// data read on an ArrayList so that we can printed the
			// content of the excel to the console.
			//
			Iterator rows = sheet.rowIterator();
			while (rows.hasNext()) {
				HSSFRow row = (HSSFRow) rows.next();

				Iterator cells = row.cellIterator();
				List data = new ArrayList();
				while (cells.hasNext()) {
					HSSFCell cell = (HSSFCell) cells.next();
					// System.out.println("Añadiendo Celda: " + cell.toString());
					data.add(cell);
				}
				sheetData.add(data);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (fis != null) {
				fis.close();
			}
		}
		showExelData(sheetData);

		// impresion de matriz
		/*
		 * for (int filas = 0; filas <fila; filas++) { for (int columna = 0; columna <
		 * 4; columna++) { if(columna<3) {
		 * 
		 * 
		 * }else { System.out.println(""); }
		 * 
		 * } }
		 */

		int matriz2[][] = new int[fila][3];

		int numeroDeDatosMatriz = 0;

		for (int filas = 0; filas < fila; filas++) {
			int validar = 0;

			for (int recorrido = 0; recorrido < numeroDeDatosMatriz; recorrido++) {
				if (validar == 1) {
					// System.out.println("si encontro");
					// break;
				}

				// System.out.print(matriz[filas][0]+"|");
				// System.out.println(matriz2[recorrido][0]+"|");

				if (matriz[filas][0] == matriz2[recorrido][0]) {
					matriz2[recorrido][1] = matriz2[recorrido][1] + matriz[filas][1];
					matriz2[recorrido][2] = matriz2[recorrido][2] + matriz[filas][2];
					validar = 1;
				}
			}

			if (validar == 0) {
				matriz2[numeroDeDatosMatriz][0] = matriz[filas][0];
				matriz2[numeroDeDatosMatriz][1] = matriz[filas][1];
				matriz2[numeroDeDatosMatriz][2] = matriz[filas][2];
				numeroDeDatosMatriz++;

			}

			if (validar == 1) {
				validar = 0;
			}

		}
		// impresion del contro general
		/*
		 * 
		 * for (int filas = 0; filas <fila; filas++) { for (int columna = 0; columna <
		 * 4; columna++) { if(columna<3) {
		 * 
		 * System.out.print(matriz2[filas][columna]+"|"); }else {
		 * System.out.println(""); }
		 * 
		 * } }
		 */
		//////////// Pedir numero
		
		int numero = Integer.parseInt( JOptionPane.showInputDialog(
			        null,"¿Qué numeros con importe arriba de ?",
			        "Petición ",
			        JOptionPane.QUESTION_MESSAGE) );
		
		
		////////
		int matrizresultados[][] = new int[fila][2];

		// filtro primer lugar

		int numeroDeresultadosPrimerLugar = 0;

		for (int filas = 0; filas < fila; filas++) {
			if (matriz2[filas][1] > numero) {
				matrizresultados[numeroDeresultadosPrimerLugar][0] = matriz2[filas][0];
				matrizresultados[numeroDeresultadosPrimerLugar][1] = matriz2[filas][1];
				numeroDeresultadosPrimerLugar++;
			}
		}

		for (int filas = 0; filas < numeroDeresultadosPrimerLugar; filas++) {
			// System.out.print(matrizresultados[filas][0]+"|");
			// System.out.println(matrizresultados[filas][1]);
		}

		//

		////////

		// System.out.println("///////////////////////////////////////////////");
		int matrizresultadosSegundo[][] = new int[fila][2];

		// filtro primer lugar

		int numeroDeresultadosSegundoLugar = 0;

		for (int filas = 0; filas < fila; filas++) {
			if (matriz2[filas][2] > numero) {
				matrizresultadosSegundo[numeroDeresultadosSegundoLugar][0] = matriz2[filas][0];
				matrizresultadosSegundo[numeroDeresultadosSegundoLugar][1] = matriz2[filas][2];
				numeroDeresultadosSegundoLugar++;
			}
		}

		for (int filas = 0; filas < numeroDeresultadosSegundoLugar; filas++) {
			// System.out.print(matrizresultadosSegundo[filas][0]+"|");
			// System.out.println(matrizresultadosSegundo[filas][1]);
		}

		//

		String textoArchivo = "";

		int numeroDeFilasCrearExcel = 0;

		for (int lugar = 0; lugar < 2; lugar++) {
			if (lugar == 0) {
				textoArchivo = "Primeros lugares";
				numeroDeFilasCrearExcel = numeroDeresultadosPrimerLugar;
			} else {
				textoArchivo = "Segundos lugares";
				numeroDeFilasCrearExcel = numeroDeresultadosSegundoLugar;
			}

			/* La ruta donde se creará el archivo */
			String rutaArchivo = System.getProperty("user.home") + "/" + textoArchivo + ".xls";
			/* Se crea el objeto de tipo File con la ruta del archivo */
			File archivoXLS = new File(rutaArchivo);
			/* Si el archivo existe se elimina */
			if (archivoXLS.exists())
				archivoXLS.delete();
			/* Se crea el archivo */
			archivoXLS.createNewFile();

			/* Se crea el libro de excel usando el objeto de tipo Workbook */
			Workbook libro = new HSSFWorkbook();
			/* Se inicializa el flujo de datos con el archivo xls */
			FileOutputStream archivo = new FileOutputStream(archivoXLS);

			/*
			 * Utilizamos la clase Sheet para crear una nueva hoja de trabajo dentro del
			 * libro que creamos anteriormente
			 */
			Sheet hoja = libro.createSheet(textoArchivo+" arriba de " + numero);

			/* Hacemos un ciclo para inicializar los valores de 10 filas de celdas */
			for (int f = 0; f < numeroDeFilasCrearExcel + 1; f++) {
				/* La clase Row nos permitirá crear las filas */
				Row fila = hoja.createRow(f);

				/* Cada fila tendrá 5 celdas de datos */
				for (int c = 0; c < 2; c++) {
					/* Creamos la celda a partir de la fila actual */
					Cell celda = fila.createCell(c);

					/* Si la fila es la número 0, estableceremos los encabezados */
					if (f == 0) {
						if (c == 0) {
							celda.setCellValue("Numero Ganador");
						} else {
							celda.setCellValue(textoArchivo+" arriba de " + numero);
						}

					} else {

						if (lugar == 0) {
							/* Si no es la primera fila establecemos un valor */
							if(c==0)
							{
								if(matrizresultados[f - 1][c]<100) 
								{
									celda.setCellValue("0" + matrizresultados[f - 1][c]);
								}else 
								{
									celda.setCellValue("" + matrizresultados[f - 1][c]);
								}
							}else 
							{
								celda.setCellValue("" + matrizresultados[f - 1][c]);
							}
						} else {
							if(c==0) 
							{
								if(matrizresultadosSegundo[f - 1][c]<100) 
								{
									celda.setCellValue("0" + matrizresultadosSegundo[f - 1][c]);
								}else 
								{
									celda.setCellValue("" + matrizresultadosSegundo[f - 1][c]);
								}
							}else 
							{
								celda.setCellValue("" + matrizresultadosSegundo[f - 1][c]);
							}
							
						}

					}
				}
			}

			/* Escribimos en el libro */
			libro.write(archivo);
			/* Cerramos el flujo de datos */
			archivo.close();
			/* Y abrimos el archivo con la clase Desktop */
			Desktop.getDesktop().open(archivoXLS);

		}

	}

	private static void showExelData(List sheetData) {
		//
		// Iterates the data and print it out to the console.
		//
		matriz = new int[sheetData.size()][3];
		fila = sheetData.size();
		for (int i = 1; i < sheetData.size(); i++) {
			// fila
			// System.out.println(sheetData.size());
			List list = (List) sheetData.get(i);
			for (int j = 0; j < list.size(); j++) {
				Cell cell = (Cell) list.get(j);
				if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
					// System.out.print(cell.getNumericCellValue());

					matriz[i][j] = (int) cell.getNumericCellValue();

				} else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
					// System.out.print(cell.getRichStringCellValue());
				} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
					// System.out.print(cell.getBooleanCellValue());
				}
				if (j < list.size() - 1) {
					// System.out.print(", ");
				}
			}
		}
	}
}