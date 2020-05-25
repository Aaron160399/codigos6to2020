/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package lecturaescrituraexcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import javax.swing.table.DefaultTableModel;
import org.apache.commons.math3.analysis.function.Ceil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 *
 * @author Aaron
 */
public class Functions {
    //Creo una constante donde almacenaré la ruta de mi archivo. Será privada
    //para que no la puedan modificar externamente
    private String PATH = "excel/prueba.xlsx";
    //Creo un libro de forma global para trbajar sobre el mismo tanto al leer
    //como escribir datos. Será privado para que no se pueda modificar externamente
    private Workbook workbook;
    
    //Con este método abriré el archivo a nivel lógica
    public DefaultTableModel loadBook() throws FileNotFoundException, IOException {
        //Declaro una bandera para saber si es la primera vuelta
        boolean firstLap = true;
        //Creo un modelo de tabla
        DefaultTableModel model = new DefaultTableModel();
        //Obtengo mi archivo a partir de una dirección URL
        InputStream inputStream = new FileInputStream(PATH);
        //"Abro" el archivo de Excel. Con este objeto ya puedo manipularlo. 
        //Inicializo el libro global que declaré
        workbook = WorkbookFactory.create(inputStream);
        //Obtengo la hoja sobre la que voy a trabajar. Comenzamos desde 0.
        XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);
        //Creo un iterador que sólo acepte filas para poder avanzar siempre que 
        //haya una celda, ya que desconoceré el número total que ya trae el archivo
        Iterator<Row> rowIterator = sheet.iterator();
        //Creo un arraylist que sólo acepte Strings en lugar de un arreglo
        //porque desconozco cuántas celdas tengo
        ArrayList<String> cellValues;
        //Mientras que tenga filas disponibles, iré dando vueltas en mi ciclo
        while (rowIterator.hasNext()) {
            //Reinicio mi arraylist
            cellValues = new ArrayList();
            //Creo una fila. Esta la obtendré de la fila que corresponde
            Row row = rowIterator.next();
            //También crearé un iterador para las columnas, porque tampoco
            //conozco el número de estas
            Iterator<Cell> cellIterator = row.cellIterator();
            //Mientras que tenga celdas disponibles, iré dando vueltas en mi ciclo
            int i = 0;
            while(cellIterator.hasNext()){
                i++;
                //Creo una celda. Esta la obtendré de la celda que corresponde
                Cell cell = cellIterator.next();
                //Obtengo el tipo de dato que tiene la celda
                try {
                    //Obtengo el valor String de la celda y lo agrego al array
                    cellValues.add(cell.getStringCellValue());
                } catch (Exception e) {
                    //También puedo obtener valores numéricos con cell.getNumericCellValue
                    double numericValue = cell.getNumericCellValue();
                    cellValues.add(String.valueOf(numericValue));
                }
            }
            //Creo un Array normal para poder agregarlo a mi tabla
            Object[] values = cellValues.toArray();
            if (firstLap) {
                //Si estoy en la primera vuelta, agrego mis encabezados
                model.setColumnIdentifiers(values);
                //Indico que ya no es la primera vuelta
                firstLap = false;
            } else {
                //Si no es la primera vuelta, agrego los valores como una fila más
                model.addRow(values);
            }
        }
        //Regreso mi modelo
        return model;
    }
    
    void writeBook(DefaultTableModel model) throws FileNotFoundException, IOException{
        //Obtengo el título que tenía la hoja
        String sheetName = workbook.getSheetAt(0).getSheetName();
        //Elimino la hoja que ya tenía
        workbook.removeSheetAt(0);
        //Creo una nueva hoja y le agrego el mismo título
        XSSFSheet sheet = (XSSFSheet) workbook.createSheet(sheetName);
        //Creo una fila
        XSSFRow rowHeaders = sheet.createRow(0);
        //Agrego los encabezados
        for (int j = 0; j < model.getColumnCount(); j++) {
            //Creo una celda
            XSSFCell cell= rowHeaders.createCell(j);
            //Agrego mi valor
            cell.setCellValue(model.getColumnName(j));
        }
        //Creo un for que dará las mismas vueltas que tiene mi modelo
        for (int i = 0; i < model.getRowCount(); i++) {
            //Creo una fila
            XSSFRow row = sheet.createRow(i+1);
            //Creo un for que dará un número de vueltas igual al número de columnas
            //que tenga mi tabla
            for (int j = 0; j < model.getColumnCount(); j++) {
                System.out.println("j: "+j);
                //Creo una celda
                XSSFCell cell= row.createCell(j);
                //Agrego mi valor
                cell.setCellValue(model.getValueAt(i, j).toString());
            }
        }
        //Creo un archivo de salida que será mi archivo de Excel
        FileOutputStream fileOutputStream = new FileOutputStream(PATH);
        //Sobreescribo mi archivo
        workbook.write(fileOutputStream);
        //Cierro mi libro
        workbook.close();
        //Cierro mi archivo de salida
        fileOutputStream.flush();
        fileOutputStream.close();
    }
}
