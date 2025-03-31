/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package conexao;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import javax.swing.JOptionPane;
/**
 *
 * @author kauam
 */
public class Conexao {
    public Connection conectar(){
        try{
           Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/ brigadeiros_tk","root","@Kaua1415");
            return con;
        }catch(SQLException sqle){
            JOptionPane.showMessageDialog(null, "Impossivel se conectar ao Banco de Dados");
            return null;
        }
    }
}
