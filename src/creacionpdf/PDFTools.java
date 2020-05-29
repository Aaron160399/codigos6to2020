/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package creacionpdf;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.File;
import java.io.FileOutputStream;
import javax.swing.table.DefaultTableModel;

/**
 *
 * @author Aaron
 */
public class PDFTools {
    /*Declaro una serie de variables globales que me ayudarán en la creadión del
    documento
    Un documento en general*/
    private Document document;
    //Una objeto del tipo File donde voy a almacenar mi documento
    private File file;
    //Un objeto que me permita modificar el PDF
    private PdfWriter pdfWriter;
    /*Creo mis valores de diseño, pero estas serán estáticas para que pueda usarlas
    desde fuera*/
    /*Una fuente para el título. Será con la fuente Helvetica, tamaño 12 en 
    negritas y color negro*/
    public static Font fTítle = new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK);
    /*Una fuente para el texto normal. Será con la fuente Helvetica, tamaño 12 en 
    normal y color negro*/
    public static Font fText = new Font(Font.FontFamily.HELVETICA, 12, Font.NORMAL, BaseColor.BLACK);
    
    
    /**
     * Este método creará una carpete en Mis Documentos y ahí guardará los PDF's
     * @param folderName - El nombre de la carpeta
     * @param documentName - El nombre del documento
     */
    private void createDirectory(String folderName, String documentName){
        //Obtengo la ruta del usuario
        String path = System.getProperty("user.home");
        //Agrego la ruta de Mis documentos
        path = path + "\\documents";
        //Creamos una carpeta donde voy a guardar mi archivo PDF
        File folder = new File(path, folderName);
        //Verifico que exista dicho folder
        if (!folder.exists()){
            //Si no existe, entonces lo creo
            folder.mkdir();
        }
        /*Creo mi archivo PDF dentro de dicha carpeta. Aquí agrego el nombre del
        documento*/
        file = new File(folder, documentName);
    }
    
    /**
     * Este método nos sirve para crear y ABRIR nuestro documento PDF
     * @param folderName - El nombre de la carpeta donde guardarpe mis PDF's
     * @param documentName - El nombre del documento
     */
    public void openDocument(String folderName, String documentName){
        //Mando a llamar mi método para crear mi carpeta y archivo
        createDirectory(folderName, documentName);
        try {
            //Inicializamos el documento y le damos tamaño de hoja CARTA
            document = new Document(PageSize.LETTER);
            /*Especifico el documento que acabo de crear y el lugar donde lo 
            voy a guardar (archivo) con esto también indico que voy a empezar a 
            escribir en él. El primer parámetro es mi documento y el segundo la
            ruta*/
            pdfWriter = PdfWriter.getInstance(document, new FileOutputStream(file));
            //Abro el documento para editarlo
            document.open();
        } catch (Exception e){
            //En caso de que haya algún error, lo guardamos en el log
            System.out.println("Error while opening the document:" +e);
        }
    }
    
    /**
     * Este método permite agregar, si así lo deseamos, los siguientes 
     * metadatos al documento
     * @param title - Título del documento
     * @param subject - Tema o asunto del documento
     * @param author - Autor del documento
     */
    public void addMetadata(String title, String subject, String author){
        //Agrego título a mi documento
        document.addTitle(title);
        //Agrego tema
        document.addSubject(subject);
        //Agrego autor
        document.addAuthor(author);
    }
    
    /**
     * Este método sirve para agregar un párrafo nuevo a mi documento
     * @param text - El párrafo a agregar
     * @param font - El estilo de fuente que usará mi párrafo
     * @param alignment - La alineación del texto. Se obtiene de la clase Paragraph
     */
    public void addParagraph(String text, Font font, int alignment){
        try{
            //Creo un nuevo párrafo
            Paragraph paragraph = new Paragraph(text, font);
            //Agrego el espacio de separación por encima de este párrafo
            paragraph.setSpacingBefore(5);
            //Agrego el espacio de separación por debajo de este párrafo
            paragraph.setSpacingAfter(5);
            //Agrego mi alineación
            paragraph.setAlignment(alignment);
            //Agrego mi párrafo al documento
            document.add(paragraph);
        } catch (Exception e){
            //En caso de que haya algún error, lo guardamos en el log
            System.out.println("Error while adding a Paragraph: "+e);
        }
    }
    
    /**
     * Este método me permitirá agregar una tabla a mi documento
     * @param model - El modelo de tabla de donde sacaré la información
     * @param font - La fuente que tendrán mis elementos
     */
    public void addTable(DefaultTableModel model, Font font) throws DocumentException{
        //Creo un nuevo párrafo
        Paragraph paragraph = new Paragraph();
        //Agrego mi fuente
        paragraph.setFont(font);
        //Creo una tabla. Me pide el número de columnas, en este caso son 3
        PdfPTable table = new PdfPTable(model.getColumnCount());
        //Le indico que ocupe el 100% del espacio disponible
        table.setWidthPercentage(100);
        //Agrego mis encabezados
        addTableHeaders(model, table);
        /*Crearé un doble for que recorrerá el contenido de mi tabla por completo.
        El primer for recorre las filas y el segundo la scolumnas*/
        for (int i = 0; i < model.getRowCount(); i++) {
            for (int j = 0; j < model.getColumnCount(); j++) {
                /*Obtengo el valor que irá en la celda, que es el contenido en las
                celdas de mi tabla*/
                String value = model.getValueAt(i, j).toString();
                /*Agrego una celda. Cada una con el valor contenido en la celda 
                correspondiente al modelo, no se combinará con nunguna otra 
                columna o fila, tendrá un borde sencillo, su alineación
                horizontal será a la derecha y tendrá la fuente definida por el usuario*/
                table.addCell(createCell(value, 1, 1, Rectangle.BOX, Element.ALIGN_LEFT, font));
            }
        }
        //Agrego la tabla a mi documento
        document.add(table);
    }
    
    /**
     * Este método agregará los encabezados a mi tabla
     * @param model - El modelo de tabla de donde sacaré los encabezados
     * @param pdfPTable - La tabla en el documento a la que los agregaré
     */
    private void addTableHeaders(DefaultTableModel model, PdfPTable pdfPTable){
        /*Creo un ciclo que dará n vueltas, donde n es el número de columnas que 
        tiene mi tabla*/
        for (int i = 0; i < model.getColumnCount(); i++) {
            /*Agrego una celda. Cada una con el nombre de mis encabezados. Cada 
            celda tiene como texto el nombre del encabezado, no se combinará con
            nunguna otra columna o fila, tendrá un borde sencillo, su alineación
            horizontal será al centro y tendrá la fuente de los títulos*/
            pdfPTable.addCell(createCell(model.getColumnName(i), 1, 1, Rectangle.BOX, Element.ALIGN_CENTER, fTítle));
        }
    }
    
    /**
     * Este método permite crear una celda con texto y con un formato específico
     * @param text - El texto que irá dentro de la celda
     * @param colspan - Cuántas columnas se quieren combinar
     * @param rowspan - Cuántas filas se quieren combinar
     * @param border - Qué diseño quiero que tenga el borde
     * @return - La celda formateada
     */
    private PdfPCell createCell(String text, int colspan, int rowspan, int border, int alignment, Font font){
        //Creo una celda
        PdfPCell cell = new PdfPCell(new Phrase(text, font));
        //Agrego cuántas columnas quiero que se combinen
        cell.setColspan(colspan);
        //Agrego cuántas filas quiero que se combinen
        cell.setRowspan(rowspan);
        //Agrego la alineación vertical dentro de la celda
        cell.setVerticalAlignment(Element.ALIGN_MIDDLE);
        //Agrego la alineación horizontal dentro de la celda
        cell.setHorizontalAlignment(alignment);
        //Agrego borde
        cell.setBorder(border);
        //Regreso la celda
        return cell;
    }
    
    /**
     * Este método permite cerrar el documento sobre el que estoy trabajando*/
    public void closeDocument(){
        //Cierro mi documento
        document.close();
    }
}
