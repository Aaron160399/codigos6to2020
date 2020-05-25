/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package auxiliaries;

import java.awt.Color;
import java.awt.Component;
import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;

/**
 *
 * @author Aaron
 */

//Debemos heredar de la clase DefaultTableCellRenderer
public class MyJTableCellRenderer extends DefaultTableCellRenderer{

    //Debemos utilizar el siguiente método (utilicen el predictor de código)
    @Override
    public Component getTableCellRendererComponent(JTable jtable, Object o, boolean bln, boolean bln1, int i, int i1) {
        //Preguntamos si el elemento que se encuentra en la celda el del tipo JLabel
        if (o instanceof JLabel) {
            //Si es así, simplemente lo devolvemos
            JLabel jLabel = (JLabel)o;
            return jLabel;
        }
        //Creo una celda
        Component cell = super.getTableCellRendererComponent(jtable, o, bln, bln1, i, i1);
        //Agrego color de fondo
        cell.setBackground(Color.yellow);
        return super.getTableCellRendererComponent(jtable, o, bln, bln1, i, i1); //To change body of generated methods, choose Tools | Templates.
    }
    
}
