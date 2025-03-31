package TelasFinacas;

import classesFinacas.DatasRelatorio;
import conexao.Conexao;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.*;
import java.text.SimpleDateFormat;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

public class RelatoriosDao {
    private Conexao conexao;
    private Connection con;
    private PreparedStatement stmt;
    private ResultSet rs;

    public RelatoriosDao() {
        this.conexao = new Conexao();
        con = this.conexao.conectar();
    }

    public void gerarRelatorioVendas(DatasRelatorio datas) {
        String sql = "SELECT "
            + "v.id AS venda_id, "
            + "v.data AS data_venda, "
            + "GROUP_CONCAT(CONCAT(p.nome, ' - ', pv.sabor, ' (', pv.qtd, ')') SEPARATOR ', ') AS produtos_comprados, "
            + "v.formaPagamento AS forma_pagamento, "
            + "v.desconto AS desconto, "
            + "SUM(pv.qtd * p.valorUnt) - v.desconto AS valor_total_com_desconto "
            + "FROM vendas v "
            + "JOIN produtovenda pv ON v.id = pv.idVenda "
            + "JOIN produto p ON pv.idProduto = p.id "
            + "WHERE v.data BETWEEN ? AND ? "
            + "GROUP BY v.id, v.data, v.formaPagamento, v.desconto;";
        
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");

        try {

            stmt = con.prepareStatement(sql);
            stmt.setString(1, sdf.format(datas.getData1()));
            stmt.setString(2, sdf.format(datas.getData2()));
            rs = stmt.executeQuery();


            FileInputStream fileInputStream = new FileInputStream("C:\\Users\\kauam\\Documents\\Modelo_Relatorio_Venda.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream); 
            Sheet sheet = workbook.getSheetAt(0); 


            int rowNum = 1; 

            while (rs.next()) {
                Row row = sheet.createRow(rowNum++); 
                
                row.createCell(0).setCellValue(rs.getString("venda_id")); 
                row.createCell(1).setCellValue(rs.getString("data_venda")); 
                row.createCell(2).setCellValue(rs.getString("produtos_comprados")); 
                row.createCell(3).setCellValue(rs.getString("forma_pagamento")); 
                row.createCell(4).setCellValue(rs.getDouble("desconto"));
                row.createCell(5).setCellValue(rs.getDouble("valor_total_com_desconto"));
            }

            String caminhoRelatorio = "C:\\Users\\kauam\\Documents\\relatorio_vendas_com_modelo.xlsx";

            FileOutputStream fileOut = new FileOutputStream(caminhoRelatorio);
            workbook.write(fileOut);
            fileOut.close(); 
            System.out.println("Relatório gerado com sucesso usando o modelo!");

            workbook.close();
            fileInputStream.close();

            abrirArquivoExcel(caminhoRelatorio);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // Fechando os recursos
            try {
                if (rs != null) rs.close();
                if (stmt != null) stmt.close();
                if (con != null) con.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }
    
    public void gerarRelatorioGastos(DatasRelatorio datas) {
    String sql = "SELECT data, descricao, qtd, formaPagamento, valor FROM gastos WHERE data BETWEEN ? AND ?;";
    SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");

    try {
        stmt = con.prepareStatement(sql);
        stmt.setString(1, sdf.format(datas.getData1()));
        stmt.setString(2, sdf.format(datas.getData2()));

        rs = stmt.executeQuery();

        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\kauam\\Documents\\Modelo_Relatorio_Gastos.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream); 
        Sheet sheet = workbook.getSheetAt(0); 

        int rowNum = 1; 

        while (rs.next()) {
            Row row = sheet.createRow(rowNum++); 

            row.createCell(0).setCellValue(rs.getString("data")); 
            row.createCell(1).setCellValue(rs.getString("descricao")); 
            row.createCell(2).setCellValue(rs.getInt("qtd")); 
            row.createCell(3).setCellValue(rs.getString("formaPagamento"));
            row.createCell(4).setCellValue(rs.getDouble("valor"));
        }

        String caminhoRelatorio = "C:\\Users\\kauam\\Documents\\relatorio_Gastos_com_modelo.xlsx";


        try (FileOutputStream fileOut = new FileOutputStream(caminhoRelatorio)) {
            workbook.write(fileOut);
        }

        System.out.println("Relatório gerado com sucesso usando o modelo!");

        workbook.close();
        fileInputStream.close();


        abrirArquivoExcel(caminhoRelatorio);

    } catch (Exception e) {
        e.printStackTrace(); 
    } finally {
        try {
            if (rs != null) rs.close();
            if (stmt != null) stmt.close();
            if (con != null) con.close();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }
}


    private void abrirArquivoExcel(String caminho) {
        try {
            File arquivoExcel = new File(caminho);
            if (arquivoExcel.exists()) {
                Desktop.getDesktop().open(arquivoExcel);
            } else {
                System.out.println("O arquivo não foi encontrado!");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
